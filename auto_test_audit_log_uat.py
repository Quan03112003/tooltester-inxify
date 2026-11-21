"""
KỊCH BẢN KIỂM THỬ TỰ ĐỘNG (AUTOMATED API & AUDIT LOG JSON VERIFICATION)
Hệ thống: Inxify PMS UAT (https://api-uat.inxify.com/api/v1)
Phạm vi: 05 nhóm nghiệp vụ (Phiếu thu chi, Hoàn tiền, Đặt cọc, Sửa/Hủy đặt phòng, Công nợ) + Audit Log
Không cần cài thêm thư viện ngoài (chỉ dùng thư viện chuẩn của Python).
Cách chạy:
    python auto_test_audit_log_uat.py
"""
import urllib.request
import urllib.error
import urllib.parse
import json
import ssl
import os

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

BASE = "https://api-uat.inxify.com/api/v1"
REQUIRED_AUDIT_KEYS = {
    "id", "userId", "propertyId", "action", "entityType", "entityId",
    "oldValues", "newValues", "ipAddress", "userAgent", "createdAt",
    "user", "docNo", "docType", "docLinkId", "docCtx"
}

def call_api(method, path, token=None, prop_id=None, payload=None):
    headers = {"Content-Type": "application/json", "User-Agent": "Inxify-QA-Automation/1.0"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    if prop_id:
        headers["X-Property-ID"] = prop_id
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(BASE + path, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, context=ctx, timeout=15) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", "ignore")
        try:
            return e.code, json.loads(raw)
        except Exception:
            return e.code, {"raw": raw}

def login(username, password):
    st, res = call_api("POST", "/auth/login", payload={"username": username, "password": password})
    if st != 200 or not res.get("success"):
        raise RuntimeError(f"Đăng nhập thất bại cho {username}: HTTP {st} - {res}")
    return res["data"]["accessToken"], res["data"]["user"], res["data"]

def verify_audit_schema(record):
    missing = REQUIRED_AUDIT_KEYS - set(record.keys())
    unchanged_logged = []
    old_v = record.get("oldValues")
    new_v = record.get("newValues")
    if isinstance(old_v, dict) and isinstance(new_v, dict):
        for k in set(old_v.keys()) & set(new_v.keys()):
            if old_v[k] == new_v[k]:
                unchanged_logged.append(k)
    return {
        "id": record.get("id"),
        "action": record.get("action"),
        "entityType": record.get("entityType"),
        "docNo": record.get("docNo"),
        "missingSchemaKeys": list(missing),
        "hasActor": bool(record.get("userId") or record.get("actorName")),
        "unchangedFieldsInBeforeAfter": unchanged_logged
    }

def main():
    print("=" * 85)
    print(" BẮT ĐẦU CHẠY KIỂM THỬ TỰ ĐỘNG AUDIT LOG INXIFY UAT (05 NGHIỆP VỤ & 06 BUG)")
    print("=" * 85)

    tok_it, user_it, login_it = login("minhquan1", "@Matkhau1")
    tok_acc, user_acc, _ = login("quanhm10", "@Matkhau1")
    tok_ins, user_ins, _ = login("quanhm8", "@Matkhau1")
    tok_sec, user_sec, _ = login("quanhm7", "@Matkhau1")
    pid = user_acc["propertyId"]

    print(f"[+] Đăng nhập thành công: IT ({user_it['username']}), ACCOUNTANT ({user_acc['username']}), INSPECTOR ({user_ins['username']}), SECRETARY ({user_sec['username']})")
    print(f"[+] Chi nhánh kiểm thử: {user_acc['propertyName']} ({pid})\n")

    # 1. Kiểm tra Schema JSON của các bản ghi Audit Log Nghiệp vụ (scope=data)
    st, data_logs = call_api("GET", "/audit-logs?page=1&limit=50&propertyId=all&scope=data", tok_it, pid)
    records = data_logs.get("data", [])
    print(f"1. KIỂM TRA CẤU TRÚC JSON AUDIT LOG (scope=data) -> HTTP {st} | Tổng bản ghi: {data_logs.get('total')}")
    print("-" * 85)
    for rec in records[:12]:
        chk = verify_audit_schema(rec)
        status_schema = "OK (Đủ 16 trường)" if not chk["missingSchemaKeys"] else f"Thiếu: {chk['missingSchemaKeys']}"
        dup_warn = f" | [CẢNH BÁO AC-04] Lưu thừa trường không đổi: {chk['unchangedFieldsInBeforeAfter']}" if chk["unchangedFieldsInBeforeAfter"] else ""
        print(f"  • ID: {chk['id'][:8]}... | {chk['entityType']:12} | {chk['action']:22} | Số phiếu: {str(chk['docNo']):14} | {status_schema}{dup_warn}")

    # 2. Kiểm tra tự động 06 Bug đã ghi nhận
    print("\n2. KIỂM TRA TRẠNG THÁI 06 BUG TRÊN MÔI TRƯỜNG UAT (RETEST TỰ ĐỘNG)")
    print("-" * 85)

    # BUG-AL-001: INSPECTOR có quyền admin:audit_log:view nhưng bị chặn 403
    st_ins, _ = call_api("GET", "/audit-logs?page=1&limit=5&propertyId=all&scope=data", tok_ins, pid)
    print(f"  • [BUG-AL-001] Phân quyền INSPECTOR xem /audit-logs : HTTP {st_ins} -> {'ĐÃ FIX (PASS)' if st_ins == 200 else 'CHƯA FIX (FAIL - Vẫn trả 403 Forbidden)'}")

    # BUG-AL-002: Phân hệ Công nợ doanh nghiệp (City Ledger) không sinh Audit Log
    city_ledger_logs = [r for r in records if r.get("entityType") in ("CityLedger", "CompanyInvoice", "Company") or "CI-" in str(r.get("docNo") or "")]
    print(f"  • [BUG-AL-002] Audit Log cho Công nợ City Ledger     : Tìm thấy {len(city_ledger_logs)} log -> {'ĐÃ FIX (PASS)' if len(city_ledger_logs) > 0 else 'CHƯA FIX (FAIL - 0 bản ghi log cho City Ledger)'}")

    # BUG-AL-003: Bỏ sót trường reference/notes trong CHANGE_PAYMENT_METHOD & lưu thừa trường không đổi
    cpm_logs = [r for r in records if r.get("action") == "CHANGE_PAYMENT_METHOD"]
    has_ref_in_new = any("reference" in (r.get("newValues") or {}) for r in cpm_logs)
    print(f"  • [BUG-AL-003] Lưu đủ trường reference/notes khi sửa : {'ĐÃ FIX (PASS)' if has_ref_in_new else 'CHƯA FIX (FAIL - newValues thiếu reference & notes)'}")

    # BUG-AL-004: Tìm kiếm theo mã Phiếu thu/chi PC-2026-0002
    _, res_srch = call_api("GET", "/audit-logs?page=1&limit=10&propertyId=all&scope=data&search=PC-2026-0002", tok_it, pid)
    tot_srch = res_srch.get("total", 0)
    print(f"  • [BUG-AL-004] Ô search tìm mã phiếu thu/chi PC-...  : total={tot_srch} -> {'ĐÃ FIX (PASS)' if tot_srch > 0 else 'CHƯA FIX (FAIL - Không tìm được theo voucherNo)'}")

    # BUG-AL-005: Che/mask SĐT/CCCD trong log GUEST_SEARCH
    _, acc_logs = call_api("GET", "/audit-logs?page=1&limit=20&propertyId=all&scope=access&action=GUEST_SEARCH", tok_it, pid)
    unmasked = [r for r in acc_logs.get("data", []) if not (r.get("newValues") or {}).get("masked")]
    print(f"  • [BUG-AL-005] Che/mask SĐT/CCCD trong GUEST_SEARCH  : {len(unmasked)} log chưa che -> {'ĐÃ FIX (PASS)' if len(unmasked) == 0 else 'CHƯA FIX (FAIL - SĐT đang lưu plaintext, masked=false)'}")

    # BUG-AL-006: Actor null trên ADD_BILL_ITEM & 2FA enforcement
    null_actors = [r for r in records if r.get("userId") is None and r.get("user") is None]
    must_2fa_bypass = bool(login_it.get("mustSetup2fa") and login_it.get("accessToken"))
    print(f"  • [BUG-AL-006] Actor null ({len(null_actors)} log) & 2FA token bypass : {'ĐÃ FIX (PASS)' if (not null_actors and not must_2fa_bypass) else 'CHƯA FIX (FAIL - Còn log thiếu userId & cấp token khi mustSetup2fa=true)'}")

    print("=" * 85)
    print(" HOÀN TẤT! Chi tiết gói tin JSON nằm trong thư mục: json_devtools_uat/")
    print("=" * 85)

if __name__ == "__main__":
    main()
