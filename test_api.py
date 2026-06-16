import requests
import json
import sys

BASE_URL = "http://127.0.0.1:5000"

def test_flow():
    log_lines = []
    def log(msg):
        log_lines.append(msg)
        # Only print ASCII parts to stdout to prevent CP1252 errors
        ascii_msg = msg.encode('ascii', errors='replace').decode('ascii')
        print(ascii_msg)
        
    log("=== STARTING API VERIFICATION ===")
    
    # 1. Test GET /api/leads (Initial state)
    log("1. Testing GET /api/leads...")
    r = requests.get(f"{BASE_URL}/api/leads")
    if r.status_code != 200:
        log(f"FAILED: status code {r.status_code}")
        return False
    leads = r.json()
    log(f"SUCCESS: Fetched {len(leads)} leads.")
    
    # 2. Test POST /api/score (Rule-based scoring fallback)
    log("2. Testing POST /api/score (Rule-based)...")
    r = requests.post(f"{BASE_URL}/api/score", json={})
    if r.status_code != 200:
        log(f"FAILED: status code {r.status_code}")
        return False
    scored_leads = r.json()['data']
    log("SUCCESS: Scoring endpoint returned 200 OK.")
    
    # Verify classifications
    log("Verifying Lead Classifications:")
    expected = {
        1: "VIP",
        2: "VIP",
        3: "VIP",
        4: "Trung bình",
        5: "Trung bình",
        6: "Rác",
        7: "Rác",
        8: "Rác",
        9: "Rác"
    }
    
    all_ok = True
    for lead in scored_leads:
        lid = lead['id']
        classification = lead['classification']
        score = lead['score']
        expected_class = expected[lid]
        ok = (classification == expected_class)
        status_text = "PASS" if ok else "FAIL"
        log(f"Lead {lid:02d}: Score={score:3d}, Class={classification:10} (Expected={expected_class}) -> {status_text}")
        if not ok:
            all_ok = False
            
    if not all_ok:
        log("Scoring verification failed!")
        return False
    log("Scoring verification PASSED!")

    # 3. Test POST /api/leads/update (Human-in-the-loop)
    log("3. Testing POST /api/leads/update...")
    update_payload = {
        "id": 4,
        "name": "Le Hoang Nam (Approved)",
        "status": "Đã duyệt",
        "reviewer_notes": "Da lien he, khach dong y tham quan nha mau vao thu Bay."
    }
    r = requests.post(f"{BASE_URL}/api/leads/update", json=update_payload)
    if r.status_code != 200:
        log(f"FAILED: status code {r.status_code}")
        return False
    updated_leads = r.json()['data']
    
    # Find updated lead
    lead4 = next(l for l in updated_leads if l['id'] == 4)
    log("Lead 4 updated status:")
    log(f"Status={lead4['status']}, ReviewerNotes={lead4['reviewer_notes']}")
    
    if lead4['status'] == "Đã duyệt" and "Le Hoang Nam (Approved)" in lead4['name']:
        log("Human-in-the-loop update: PASSED!")
    else:
        log("Human-in-the-loop update: FAILED!")
        return False

    # 4. Test GET /api/export (Excel generation)
    log("4. Testing GET /api/export...")
    r = requests.get(f"{BASE_URL}/api/export")
    if r.status_code != 200:
        log(f"FAILED: status code {r.status_code}")
        return False
    content_len = len(r.content)
    log(f"SUCCESS: Exported Excel file. Size = {content_len} bytes.")
    if content_len > 1000:
        log("Excel export: PASSED!")
    else:
        log("Excel export: FAILED! Size is too small.")
        return False

    log("=== ALL TESTS PASSED! ===")
    
    # Write full log to file
    with open("test_report.txt", "w", encoding="utf-8") as f:
        f.write("\n".join(log_lines))
    print("Full report saved to test_report.txt")
    return True

if __name__ == "__main__":
    success = test_flow()
    sys.exit(0 if success else 1)
