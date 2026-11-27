"""Вход и права на заявки. Нужен PostgreSQL: TEST_DATABASE_URL=postgresql://user:pass@host/db pytest tests"""
import importlib
import json
import os
import sys

import pytest

DB = os.environ.get("TEST_DATABASE_URL")
pytestmark = pytest.mark.skipif(not DB, reason="TEST_DATABASE_URL не задан")

ROOT = os.path.join(os.path.dirname(__file__), "..")


@pytest.fixture()
def mod(monkeypatch):
    monkeypatch.setenv("DATABASE_URL", DB)
    monkeypatch.setenv("SECRET_KEY", "test-secret")
    monkeypatch.setenv("BRANCH_PASSWORD", "branch-pass-1")
    monkeypatch.setenv("USER_PASSWORDS", json.dumps({"admin": "admin-pass-1", "ANNA_V": "region-pass-1"}))
    monkeypatch.delenv("RAILWAY_ENVIRONMENT", raising=False)
    monkeypatch.syspath_prepend(ROOT)
    sys.modules.pop("app", None)
    m = importlib.import_module("app")
    m.app.config["TESTING"] = True
    conn = m.get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM order_items")
    cur.execute("DELETE FROM orders")
    ids = {}
    for br in ("ECO PARK", "MAGIC CITY"):
        cur.execute("INSERT INTO orders(branch,responsible,date,status,filename) VALUES(%s,'x','2026-01-01','Новая','f.xlsx') RETURNING id", (br,))
        ids[br] = cur.fetchone()["id"]
    conn.commit()
    cur.close()
    m.release_db(conn)
    m._login_attempts.clear()
    m.ids = ids
    return m


def login(c, u, p):
    return c.post("/api/login", json={"username": u, "password": p})


def test_old_hardcoded_passwords_do_not_work(mod):
    c = mod.app.test_client()
    for u, p in (("admin", "123456"), ("ECO PARK", "LI-NING1"), ("DMITRIY", "Dm!LN2026$"), ("SKLAD", "sklad2026")):
        assert login(c, u, p).status_code == 401


def test_login_with_env_passwords(mod):
    c = mod.app.test_client()
    assert login(c, "admin", "admin-pass-1").status_code == 200
    assert login(mod.app.test_client(), "ECO PARK", "branch-pass-1").status_code == 200
    assert login(mod.app.test_client(), "ECO PARK", "wrong").status_code == 401


def test_user_without_password_in_env_cannot_login(mod):
    assert login(mod.app.test_client(), "DMITRIY", "").status_code == 401
    assert login(mod.app.test_client(), "DMITRIY", "anything").status_code == 401


def test_lockout_per_username_even_from_other_ip(mod):
    for i in range(10):
        c = mod.app.test_client()
        c.environ_base["REMOTE_ADDR"] = f"10.0.0.{i}"
        assert login(c, "admin", "bad").status_code == 401
    c = mod.app.test_client()
    c.environ_base["REMOTE_ADDR"] = "10.9.9.9"
    assert login(c, "admin", "admin-pass-1").status_code == 429


def test_password_rotation_kills_session(mod, monkeypatch):
    c = mod.app.test_client()
    assert login(c, "admin", "admin-pass-1").status_code == 200
    assert c.get("/api/me").get_json()["logged_in"] is True
    monkeypatch.setenv("USER_PASSWORDS", json.dumps({"admin": "admin-pass-2"}))
    assert c.get("/api/me").get_json()["logged_in"] is False


def test_branch_user_cannot_read_other_branch_order(mod):
    c = mod.app.test_client()
    assert login(c, "ECO PARK", "branch-pass-1").status_code == 200
    own, other = mod.ids["ECO PARK"], mod.ids["MAGIC CITY"]
    assert c.get(f"/api/orders/{own}/items").status_code == 200
    assert c.get(f"/api/orders/{other}/items").status_code == 403
    assert c.get(f"/api/orders/{other}/excel").status_code == 403
    assert c.get(f"/api/download/{other}").status_code == 403
    assert c.get(f"/api/download/{own}").status_code != 403


def test_regional_limited_to_own_branches_and_admin_sees_all(mod):
    c = mod.app.test_client()
    assert login(c, "ANNA_V", "region-pass-1").status_code == 200
    assert c.get(f"/api/orders/{mod.ids['MAGIC CITY']}/items").status_code == 200   # входит в её список
    assert c.get(f"/api/orders/{mod.ids['ECO PARK']}/items").status_code == 403     # не входит
    a = mod.app.test_client()
    assert login(a, "admin", "admin-pass-1").status_code == 200
    assert a.get(f"/api/orders/{mod.ids['ECO PARK']}/items").status_code == 200
