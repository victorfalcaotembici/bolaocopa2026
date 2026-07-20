import requests, json, os, unicodedata
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

def _norm_header(s):
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    for ch in (" ", "-"):
        s = s.replace(ch, "_")
    return s

def _norm_name(s):
    # normaliza identidade p/ join: minúsculo, sem acento, sem pontos/espaços/_/-
    # une "cintia.grispo", "Cintia Grispo" e "cíntia.grispo" na mesma chave -> "cintiagrispo"
    if not s:
        return ""
    s = str(s).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    for ch in (".", " ", "_", "-"):
        s = s.replace(ch, "")
    return s

# aliases aceitos por campo (todos normalizados: minúsculo, sem acento, espaços->_)
_FIELD_ALIASES = {
    "email":     ["email", "e_mail", "mail"],
    "id_user":   ["id_user", "iduser", "id", "usuario", "user", "login"],
    "pais":      ["pais", "country"],
    "diretoria": ["diretoria", "direccion", "directorate", "area"],
    "adm_corp":  ["adm_corp", "admcorp", "corp", "corporativo", "adm"],
    "praca":     ["praca", "cidade", "ciudad", "city"],
}

def _build_col_index(header_row):
    norm = [_norm_header(h) for h in header_row]
    idx = {}
    for field, aliases in _FIELD_ALIASES.items():
        found = None
        for a in aliases:
            if a in norm:
                found = norm.index(a)
                break
        idx[field] = found
    return idx

# posições de fallback caso o cabeçalho não seja reconhecido
_FALLBACK_POS = {"email":0, "id_user":1, "pais":2, "diretoria":3, "adm_corp":4, "praca":5}

def load_users_from_xlsx():
    wb = openpyxl.load_workbook("Base_hdc_copa.xlsx")
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return {}

    col = _build_col_index(rows[0])
    # se algum campo essencial não foi achado pelo nome, cai pro fallback posicional
    for f, pos in _FALLBACK_POS.items():
        if col.get(f) is None:
            col[f] = pos

    def get(row, field):
        i = col[field]
        return row[i] if i is not None and i < len(row) else None

    users = {}
    for row in rows[1:]:
        if not row:
            continue
        id_user   = get(row, "id_user")
        pais      = get(row, "pais")
        diretoria = get(row, "diretoria")
        adm_corp  = get(row, "adm_corp")
        praca     = get(row, "praca")
        if not id_user:
            continue
        key = _norm_name(id_user)
        praca_val = str(praca).strip() if praca else "Corporativo"
        if praca_val.lower() == "riviera":
            praca_val = "BikeSampa"
        users[key] = {
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
    for e in standings:
        u = e.get("user", {})
        dn     = (u.get("displayName") or "").strip()
        handle = (u.get("handle") or "").strip().lower()
        # join primário: displayName (identidade corporativa pontuada) vs id_user.
        # fallback: handle (só acerta quando é igual ao id sem pontos; vanity strings não batem).
        info = users_map.get(_norm_name(dn)) or users_map.get(_norm_name(handle))
        if not info:
            unmatched.append(dn or handle)  # dotted p/ localizar na base
            continue  # fora da base = fora de TODO ranking; só registra em unmatchedHandles
        enriched.append({
            "rank":             e.get("rank"),
            "totalPoints":      e.get("totalPoints", 0),
            "prevPosition":     e.get("prevPosition", 0),
            "prevTotalPoints":  e.get("prevTotalPoints", 0),
            "predictionsCount": e.get("predictionsCount", 0),
            "exactScoreCount":  e.get("exactScoreCount", 0),
            "correctWinnerCount": e.get("correctWinnerCount", 0),
            "handle":      handle,
            "displayName": dn or handle,
            "praca":       info["praca"],
            "pais":        info["pais"],
            "diretoria":   info["diretoria"],
            "adm_corp":    info["adm_corp"],
        })
    return {
        "updatedAt": datetime.utcnow().isoformat() + "Z",
        "finishedMatchesCount": len(finished),
        "standings": enriched,
        "unmatchedHandles": unmatched
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
    print(f"OK — {n} participantes, {j} jogos encerrados")
    if output["unmatchedHandles"]:
        print(f"Sem match ({len(output['unmatchedHandles'])}): {output['unmatchedHandles'][:10]}")
