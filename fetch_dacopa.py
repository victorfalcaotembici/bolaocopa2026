import requests, json, os
from datetime import datetime
import openpyxl

FIREBASE_REFRESH_TOKEN = os.environ.get("FIREBASE_REFRESH_TOKEN")
FIREBASE_API_KEY = os.environ.get("FIREBASE_API_KEY")
GROUP_CODE = "F3XVPHFW"

def get_fresh_jwt():
    url = f"https://securetoken.googleapis.com/v1/token?key={FIREBASE_API_KEY}"
    r = requests.post(url, json={"grant_type":"refresh_token","refresh_token":FIREBASE_REFRESH_TOKEN}, timeout=15)
    r.raise_for_status()
    return r.json()["id_token"]

def fetch_leaderboard(jwt):
    url = f"https://api.dacopa.com/groups/{GROUP_CODE}/leaderboard"
    headers = {
        "Authorization": f"Bearer {jwt}",
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Origin": "https://app.dacopa.com",
        "X-Client-Capabilities": "predictions-v2"
    }
    r = requests.get(url, headers=headers, timeout=15)
    r.raise_for_status()
    return r.json()

def load_users_from_xlsx():
    wb = openpyxl.load_workbook("Base_hdc_copa.xlsx")
    ws = wb.active
    users = {}
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if i == 1:
            continue
        email, id_user, pais, diretoria, adm_corp, praca = row if row else (None,)*6
        if not id_user:
            continue
        key = str(id_user).strip().lower()
        praca_val = str(praca).strip() if praca else "Corporativo"
        if praca_val.lower() == "riviera":
            praca_val = "BikeSampa"
        users[key] = {
            "email": str(email).strip() if email else "",
            "praca": praca_val,
            "pais": str(pais).strip() if pais else "Brasil",
            "diretoria": str(diretoria).strip() if diretoria else "CBO",
            "adm_corp": str(adm_corp).strip() if adm_corp else "Operação"
        }
    return users

def build_output(raw, users_map):
    standings = raw.get("standings", [])
    finished  = raw.get("finishedMatches", [])
    enriched, unmatched = [], []

    # Build set of displayNames that are enrolled (present in leaderboard)
    enrolled_names = set()
    for e in standings:
        u = e.get("user", {})
        display_key = (u.get("displayName") or "").strip().lower()
        enrolled_names.add(display_key)

    for e in standings:
        u = e.get("user", {})
        handle = (u.get("handle") or "").strip().lower()
        display_key = (u.get("displayName") or "").strip().lower()
        info = users_map.get(display_key) or users_map.get(handle)
        if not info:
            unmatched.append(handle)
            continue  # não aloca em nenhuma categoria
        enriched.append({
            "rank":             e.get("rank"),
            "totalPoints":      e.get("totalPoints", 0),
            "prevPosition":     e.get("prevPosition", 0),
            "prevTotalPoints":  e.get("prevTotalPoints", 0),
            "predictionsCount": e.get("predictionsCount", 0),
            "exactScoreCount":  e.get("exactScoreCount", 0),
            "correctWinnerCount": e.get("correctWinnerCount", 0),
            "handle":      handle,
            "displayName": u.get("displayName", handle),
            "praca":       info["praca"],
            "pais":        info["pais"],
            "diretoria":   info["diretoria"],
            "adm_corp":    info["adm_corp"],
        })

    # Build enrollment list: every person in the Excel base with their status
    enrollment = []
    for handle, info in users_map.items():
        enrollment.append({
            "handle":   handle,
            "email":    info.get("email", ""),
            "praca":    info["praca"],
            "pais":     info["pais"],
            "diretoria": info["diretoria"],
            "adm_corp": info["adm_corp"],
            "inscrito": handle in enrolled_names,
        })
    # Sort: praça asc, then uninscribed first, then handle
    enrollment.sort(key=lambda x: (x["praca"], not x["inscrito"], x["handle"]))

    return {
        "updatedAt": datetime.utcnow().isoformat() + "Z",
        "finishedMatchesCount": len(finished),
        "standings": enriched,
        "unmatchedHandles": unmatched,
        "enrollment": enrollment,
    }

if __name__ == "__main__":
    print("Carregando base do XLSX...")
    users = load_users_from_xlsx()
    print(f"Carregado {len(users)} usuários")
    
    print("Renovando JWT...")
    jwt = get_fresh_jwt()
    print("Buscando leaderboard...")
    raw = fetch_leaderboard(jwt)
    output = build_output(raw, users)
    with open("data.json", "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    n = len(output["standings"])
    j = output["finishedMatchesCount"]
    e = len(output["enrollment"])
    ins = sum(1 for x in output["enrollment"] if x["inscrito"])
    print(f"OK — {n} no ranking, {j} jogos encerrados")
    print(f"Inscrições: {ins}/{e} ({100*ins//e if e else 0}%)")
    if output["unmatchedHandles"]:
        print(f"Sem match ({len(output['unmatchedHandles'])}): {output['unmatchedHandles'][:10]}")
