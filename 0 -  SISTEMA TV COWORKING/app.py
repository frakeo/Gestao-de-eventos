# app.py — versão reforçada com suporte a hora_inicio / hora_fim
import os
import tempfile
import logging
import re
import pandas as pd
from flask import Flask, render_template, request, jsonify

app = Flask(__name__)

# CONFIGURAÇÕES
EXCEL_PATH = r"A:\0 - SISTEMA TV COWORKING\eventos.xlsx"
# Agora usamos hora_inicio e hora_fim em vez de horario
COLUMNS = ["ID", "evento", "organizador", "data", "hora_inicio", "hora_fim", "publico", "sala"]

# Logger
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s: %(message)s")
logger = logging.getLogger("eventos_app")

# Helpers
def ensure_excel_dir():
    folder = os.path.dirname(EXCEL_PATH) or "."
    os.makedirs(folder, exist_ok=True)

def _try_read_excel(path):
    """
    Tenta ler o Excel com várias engines. Lança exceção caso falhe.
    """
    try:
        return pd.read_excel(path, dtype=str, engine="openpyxl")
    except Exception as e_openpyxl:
        logger.warning("Falha ao ler com openpyxl: %s", e_openpyxl)
        try:
            return pd.read_excel(path, dtype=str)
        except Exception as e_fallback:
            logger.exception("Falha ao ler Excel com fallback.")
            raise

_time_re = re.compile(r"^([01]\d|2[0-3]):[0-5]\d$")

def validar_hora(h):
    """Retorna True se h estiver no formato HH:MM."""
    if not isinstance(h, str):
        return False
    return bool(_time_re.match(h.strip()))

def parse_horario_to_inicio_fim(horario_raw):
    """
    Tenta decompor um campo antigo 'horario' em (hora_inicio, hora_fim).
    Aceita formatos como:
      - "08:00 às 12:00"
      - "08:00-12:00"
      - "08:00 to 12:00"
      - "08:00" (retorna hora_inicio="08:00", hora_fim="")
    Se não for possível, retorna ("","").
    """
    if not horario_raw or not isinstance(horario_raw, str):
        return "", ""
    s = horario_raw.strip()
    # substituir variações por separador padrão
    s = s.replace("–", "-").replace("—", "-").replace(" to ", "-").replace(" às ", "-").replace(" as ", "-")
    parts = [p.strip() for p in re.split(r"[-–—]", s) if p.strip()]
    if len(parts) == 0:
        return "", ""
    if len(parts) == 1:
        part = parts[0]
        if validar_hora(part):
            return part, ""
        return "", ""
    # len >= 2
    inicio = parts[0]
    fim = parts[1]
    if validar_hora(inicio) and validar_hora(fim):
        return inicio, fim
    # fallback: try to find times inside string
    found = re.findall(r"([01]\d|2[0-3]):[0-5]\d", horario_raw)
    if len(found) >= 2:
        return found[0], found[1]
    if len(found) == 1:
        return found[0], ""
    return "", ""

def carregar_excel():
    """
    Retorna DataFrame com colunas garantidas (IDs como string).
    Se arquivo não existir, cria DataFrame vazio com colunas.
    Se arquivo contém coluna 'horario' antiga, tenta migrar para hora_inicio/hora_fim.
    """
    ensure_excel_dir()

    if os.path.exists(EXCEL_PATH):
        try:
            df = _try_read_excel(EXCEL_PATH)
        except Exception as e:
            logger.exception("Erro ao ler Excel em %s", EXCEL_PATH)
            raise
    else:
        df = pd.DataFrame(columns=COLUMNS)

    # Normalizar colunas: garantir que as colunas novas existam
    # Se arquivo antigo tem 'horario' e não tem 'hora_inicio'/'hora_fim', migramos.
    if "horario" in df.columns and ("hora_inicio" not in df.columns or "hora_fim" not in df.columns):
        logger.info("Coluna 'horario' detectada — tentando migrar para 'hora_inicio'/'hora_fim'.")
        df["hora_inicio"] = ""
        df["hora_fim"] = ""
        for idx, row in df.iterrows():
            raw = row.get("horario", "")
            hi, hf = parse_horario_to_inicio_fim(raw)
            df.at[idx, "hora_inicio"] = hi
            df.at[idx, "hora_fim"] = hf
        # opcional: você poderia remover 'horario' daqui — mas vamos manter até salvar, onde salvamos apenas COLUMNS

    # Garantir todas as colunas esperadas existam
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = ""

    # Reordenar e preencher NaN
    df = df[COLUMNS].fillna("")
    # Garantir ID como string
    df["ID"] = df["ID"].astype(str)

    return df

def salvar_excel_atomico(df):
    """
    Salva o DataFrame de forma atômica (temp file + replace).
    Garante ordem das colunas conforme COLUMNS.
    """
    ensure_excel_dir()
    # garantir colunas na ordem
    df = df[COLUMNS]

    dir_dest = os.path.dirname(EXCEL_PATH) or "."
    fd = None
    tmp_path = None
    try:
        fd, tmp_path = tempfile.mkstemp(suffix=".xlsx", dir=dir_dest)
        os.close(fd)
        try:
            with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False)
        except Exception as e_writer:
            logger.warning("Escrita com openpyxl falhou: %s — tentando sem engine explícita", e_writer)
            df.to_excel(tmp_path, index=False)

        os.replace(tmp_path, EXCEL_PATH)
        logger.info("Arquivo salvo com sucesso em %s", EXCEL_PATH)
    except PermissionError:
        logger.exception("PermissionError ao salvar Excel — provavelmente o arquivo está aberto no Excel.")
        if tmp_path and os.path.exists(tmp_path):
            try: os.remove(tmp_path)
            except Exception: pass
        raise
    except Exception:
        logger.exception("Erro ao salvar Excel.")
        if tmp_path and os.path.exists(tmp_path):
            try: os.remove(tmp_path)
            except Exception: pass
        raise

def gerar_proximo_id(df):
    if df.empty:
        return "1"
    numeric_ids = []
    for v in df["ID"].tolist():
        try:
            numeric_ids.append(int(v))
        except Exception:
            pass
    if not numeric_ids:
        return "1"
    return str(max(numeric_ids) + 1)

# Rotas
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/cadastro", methods=["GET"])
def cadastro_get():
    return render_template("cadastro.html")

@app.route("/eventos", methods=["GET"])
def eventos():
    try:
        df = carregar_excel()
        # retornar registros como lista de dicts
        records = df.to_dict(orient="records")
        return jsonify(records)
    except Exception as e:
        logger.exception("Erro ao fornecer /eventos")
        return jsonify({"error": "Erro ao ler eventos", "detail": str(e)}), 500

def _extract_hours_from_request(form):
    """
    Extrai hora_inicio/hora_fim do form.
    Aceita keys: 'horaInicio', 'hora_inicio', 'horario' (antigo).
    Retorna tuple (hora_inicio, hora_fim).
    """
    h_inicio = (form.get("horaInicio") or form.get("hora_inicio") or "").strip()
    h_fim = (form.get("horaFim") or form.get("hora_fim") or "").strip()

    # If both empty but 'horario' exists, try parse
    if not h_inicio and not h_fim:
        horario_raw = (form.get("horario") or "").strip()
        if horario_raw:
            hi, hf = parse_horario_to_inicio_fim(horario_raw)
            return hi, hf
    return h_inicio, h_fim

@app.route("/cadastro", methods=["POST"])
def cadastro():
    try:
        df = carregar_excel()

        evento = request.form.get("evento", "").strip()
        organizador = request.form.get("organizador", "").strip()
        data = request.form.get("data", "").strip()
        publico = request.form.get("publico", "SRA-ES").strip()
        sala = request.form.get("sala", "").strip()

        hora_inicio, hora_fim = _extract_hours_from_request(request.form)

        # validações básicas
        if not all([evento, data, sala]):
            return jsonify({"error": "Campos obrigatórios faltando (evento, data ou sala)"}), 400
        if not hora_inicio or not hora_fim:
            return jsonify({"error": "Hora de início e término obrigatórias (hora_inicio/hora_fim)."}), 400
        if not validar_hora(hora_inicio) or not validar_hora(hora_fim):
            return jsonify({"error": "Formato de hora inválido. Use HH:MM."}), 400

        novo_id = gerar_proximo_id(df)

        novo = {
            "ID": novo_id,
            "evento": evento,
            "organizador": organizador,
            "data": data,
            "hora_inicio": hora_inicio,
            "hora_fim": hora_fim,
            "publico": publico,
            "sala": sala
        }

        df = pd.concat([df, pd.DataFrame([novo])], ignore_index=True)
        salvar_excel_atomico(df)
        return jsonify({"ok": True, "id": novo_id}), 200

    except PermissionError:
        return jsonify({"error": "Arquivo Excel está aberto ou protegido. Feche o arquivo e tente novamente."}), 423
    except Exception as e:
        logger.exception("Erro ao cadastrar evento")
        return jsonify({"error": "Erro interno no cadastro", "detail": str(e)}), 500

@app.route("/editar/<id>", methods=["POST"])
def editar(id):
    try:
        df = carregar_excel()
        id = str(id)

        if id not in df["ID"].values:
            return jsonify({"error": "ID não encontrado"}), 404

        evento = request.form.get("evento", "").strip()
        organizador = request.form.get("organizador", "").strip()
        data = request.form.get("data", "").strip()
        publico = request.form.get("publico", "SRA-ES").strip()
        sala = request.form.get("sala", "").strip()

        hora_inicio, hora_fim = _extract_hours_from_request(request.form)

        if not all([evento, data, sala]):
            return jsonify({"error": "Campos obrigatórios faltando (evento, data ou sala)"}), 400
        if not hora_inicio or not hora_fim:
            return jsonify({"error": "Hora de início e término obrigatórias (hora_inicio/hora_fim)."}), 400
        if not validar_hora(hora_inicio) or not validar_hora(hora_fim):
            return jsonify({"error": "Formato de hora inválido. Use HH:MM."}), 400

        idx = df.index[df["ID"] == id][0]

        df.at[idx, "evento"] = evento
        df.at[idx, "organizador"] = organizador
        df.at[idx, "data"] = data
        df.at[idx, "hora_inicio"] = hora_inicio
        df.at[idx, "hora_fim"] = hora_fim
        df.at[idx, "publico"] = publico
        df.at[idx, "sala"] = sala

        salvar_excel_atomico(df)
        return jsonify({"ok": True}), 200

    except PermissionError:
        return jsonify({"error": "Arquivo Excel está aberto ou protegido. Feche o arquivo e tente novamente."}), 423
    except Exception as e:
        logger.exception("Erro ao editar evento")
        return jsonify({"error": "Erro interno ao editar", "detail": str(e)}), 500

@app.route("/cancelar/<id>", methods=["POST"])
def cancelar(id):
    try:
        df = carregar_excel()
        id = str(id)

        if id not in df["ID"].values:
            return jsonify({"error": "ID não encontrado"}), 404

        df = df[df["ID"] != id]
        salvar_excel_atomico(df)
        return jsonify({"ok": True}), 200

    except PermissionError:
        return jsonify({"error": "Arquivo Excel está aberto ou protegido. Feche o arquivo e tente novamente."}), 423
    except Exception as e:
        logger.exception("Erro ao cancelar evento")
        return jsonify({"error": "Erro interno ao cancelar", "detail": str(e)}), 500

@app.route("/_debug_paths")
def debug_paths():
    return jsonify({"EXCEL_PATH": EXCEL_PATH, "exists": os.path.exists(EXCEL_PATH)})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
