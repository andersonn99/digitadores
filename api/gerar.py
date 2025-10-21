# api/gerar.py
# -*- coding: utf-8 -*-
import json
import re
from http.server import BaseHTTPRequestHandler
from pathlib import Path
import openpyxl

# Caminho da planilha no projeto Vercel:
PLANILHA = Path(__file__).parent.parent / "dados" / "GABARITO.xlsx"

# Ex.: BA032025.01.00**39367CAF  (ou ...00XX..., ou ...00dd...)
CAF_RE = re.compile(
    r"""^\s*
        (?P<UF>[A-Z]{2})
        (?P<MES>\d{2})
        (?P<ANO>\d{4})
        \.
        (?P<COD1>\d{2})
        \.
        (?P<COD2>\d{2})
        (?P<IDENT>(\*\*|XX|\d{2}))
        (?P<TAIL>\d{5})
        CAF
        \s*$""",
    re.VERBOSE | re.IGNORECASE
)

def _send_json(handler: BaseHTTPRequestHandler, status: int, payload: dict):
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    # CORS
    handler.send_header("Access-Control-Allow-Origin", "*")
    handler.send_header("Access-Control-Allow-Headers", "Content-Type")
    handler.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)

def parse_caf(caf: str):
    if not caf:
        return None
    caf_n = caf.strip().upper().replace(" ", "")
    m = CAF_RE.match(caf_n)
    if not m:
        return None
    d = m.groupdict()
    d["CAF_NORMALIZADO"] = caf_n
    return d

def ler_identificadores_unicos_ordenados(mes: str, ano: str):
    """Lê coluna F (índice 5) para o mês/ano e retorna lista de strings '00'..'99' únicas e ordenadas."""
    if not PLANILHA.exists():
        raise FileNotFoundError(f"Planilha não encontrada: {PLANILHA}")

    wb = openpyxl.load_workbook(PLANILHA, data_only=True)
    ws = wb.active

    vistos = set()
    ids = []

    # Colunas esperadas: B (2) = mês, C (3) = ano, F (6) = identificador
    for row in ws.iter_rows(min_row=2, values_only=True):
        mes_val = row[1]
        ano_val = row[2]
        ident_val = row[5] if len(row) >= 6 else None

        try:
            mes_ok = f"{int(mes_val):02d}" if mes_val is not None else None
        except Exception:
            mes_ok = None
        try:
            ano_ok = str(int(ano_val)) if ano_val is not None else None
        except Exception:
            ano_ok = None

        if mes_ok == mes and ano_ok == ano:
            # identificador deve ser número 0..99
            try:
                n = int(str(ident_val).strip())
            except Exception:
                continue
            if 0 <= n <= 99:
                cod = f"{n:02d}"
                if cod not in vistos:
                    vistos.add(cod)
                    ids.append(cod)

    ids.sort(key=lambda x: int(x))
    return ids

def gerar_combos(parsed: dict, ids_lista):
    """Monta prefixo e substitui IDENT pelos ids da planilha.
       Se IDENT já for 'dd', valida e retorna apenas 1 código (se existir no conjunto)."""
    prefixo = f"{parsed['UF']}{parsed['MES']}{parsed['ANO']}.{parsed['COD1']}.{parsed['COD2']}"
    tail = f"{parsed['TAIL']}CAF"

    ident = parsed["IDENT"]
    # Caso IDENT já seja dois dígitos:
    if re.fullmatch(r"\d{2}", ident):
        # Se o dígito existe na lista permitida, retorna 1 único; senão, lista vazia
        return [f"{prefixo}{ident}{tail}"] if ident in ids_lista else []

    # Caso IDENT seja ** ou XX -> gera para todos os ids válidos
    return [f"{prefixo}{cod}{tail}" for cod in ids_lista]

class handler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.end_headers()

    def do_POST(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(length) if length else b"{}"
            data = json.loads(raw.decode("utf-8") or "{}")
            caf_in = (data.get("caf") or "").strip()

            parsed = parse_caf(caf_in)
            if not parsed:
                return _send_json(self, 400, {"erro": "Formato inválido. Ex.: BA032025.01.00**39367CAF"})

            mes = parsed["MES"]
            ano = parsed["ANO"]

            ids = ler_identificadores_unicos_ordenados(mes, ano)
            if not ids:
                return _send_json(self, 404, {"erro": f"Nenhum identificador encontrado para {mes}/{ano}."})

            combos = gerar_combos(parsed, ids)
            if not combos:
                return _send_json(self, 404, {"erro": "IDENT já numérico e não presente no gabarito para este mês/ano."})

            return _send_json(self, 200, {"combos": combos})

        except FileNotFoundError as e:
            return _send_json(self, 500, {"erro": str(e)})
        except Exception as e:
            return _send_json(self, 500, {"erro": f"Erro interno: {e}"})
