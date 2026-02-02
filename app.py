# app.py
import os
import uuid
import json

from flask import Flask, render_template, request, redirect, url_for, session, send_file, flash
from werkzeug.utils import secure_filename

from core.step1_banco_consolidado import gerar_banco_consolidado
from core.step2_gerar_espelhos import gerar_espelhos_motoristas
from core.step3_resumos import gerar_resumos

# =========================
# CONFIG
# =========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

STORAGE_DIR = os.path.join(BASE_DIR, "storage")
UPLOADS_DIR = os.path.join(STORAGE_DIR, "uploads")
WORKSPACES_DIR = os.path.join(STORAGE_DIR, "workspaces")
DOWNLOADS_DIR = os.path.join(STORAGE_DIR, "downloads")

MODELO_PATH = os.path.join(BASE_DIR, "modelo", "modelo.xlsx")

ALLOWED_EXTENSIONS = {".xlsx"}

os.makedirs(UPLOADS_DIR, exist_ok=True)
os.makedirs(WORKSPACES_DIR, exist_ok=True)
os.makedirs(DOWNLOADS_DIR, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "troque-essa-chave-em-producao")

# (Opcional)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

# =========================
# 🔒 SENHA FIXA (ALTERE AQUI)
# =========================
APP_PASSWORD = "espelho2026"


# =========================
# HELPERS
# =========================
def _ext_ok(filename: str) -> bool:
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXTENSIONS


def _get_sid() -> str:
    sid = session.get("sid")
    if not sid:
        sid = uuid.uuid4().hex
        session["sid"] = sid
    return sid


def _ws_dir() -> str:
    sid = _get_sid()
    d = os.path.join(WORKSPACES_DIR, sid)
    os.makedirs(d, exist_ok=True)
    return d


def _dl_dir() -> str:
    sid = _get_sid()
    d = os.path.join(DOWNLOADS_DIR, sid)
    os.makedirs(d, exist_ok=True)
    return d


def _state_path() -> str:
    return os.path.join(_ws_dir(), "state.json")


def load_state() -> dict:
    p = _state_path()
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    return {
        "uploaded": False,
        "step1_done": False,
        "step2_done": False,
        "step3_done": False,
        "files": {}
    }


def save_state(state: dict) -> None:
    with open(_state_path(), "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def require_uploaded_files(state: dict):
    files = state.get("files", {})
    motoristas = files.get("motoristas")
    fechamento = files.get("fechamento")
    if not motoristas or not fechamento:
        raise ValueError("Arquivos não enviados.")
    if not os.path.exists(motoristas) or not os.path.exists(fechamento):
        raise FileNotFoundError("Arquivos enviados não encontrados no workspace.")


def is_logged_in() -> bool:
    return session.get("auth_ok") is True


# =========================
# 🔒 LOGIN ROUTES
# =========================
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        pwd = request.form.get("password", "").strip()
        if pwd == APP_PASSWORD:
            session["auth_ok"] = True
            flash("Acesso liberado ✅", "ok")
            return redirect(url_for("index"))
        flash("Senha incorreta.", "error")
        return redirect(url_for("login"))

    return render_template("login.html")


@app.route("/logout", methods=["GET"])
def logout():
    session.clear()
    flash("Você saiu.", "ok")
    return redirect(url_for("login"))


# =========================
# ROUTES
# =========================
@app.route("/", methods=["GET"])
def index():
    if not is_logged_in():
        return redirect(url_for("login"))

    state = load_state()
    return render_template("index.html", state=state)


@app.route("/upload", methods=["POST"])
def upload():
    if not is_logged_in():
        return redirect(url_for("login"))

    state = load_state()
    ws = _ws_dir()

    motoristas_file = request.files.get("motoristas")
    fechamento_file = request.files.get("fechamento")

    if not motoristas_file or not fechamento_file:
        flash("Envie os dois arquivos (motoristas e fechamento).", "error")
        return redirect(url_for("index"))

    if motoristas_file.filename == "" or fechamento_file.filename == "":
        flash("Nome de arquivo inválido.", "error")
        return redirect(url_for("index"))

    if not _ext_ok(motoristas_file.filename) or not _ext_ok(fechamento_file.filename):
        flash("Apenas arquivos .xlsx são aceitos.", "error")
        return redirect(url_for("index"))

    motoristas_name = secure_filename(motoristas_file.filename)
    fechamento_name = secure_filename(fechamento_file.filename)

    motoristas_path = os.path.join(ws, f"motoristas__{motoristas_name}")
    fechamento_path = os.path.join(ws, f"fechamento__{fechamento_name}")

    motoristas_file.save(motoristas_path)
    fechamento_file.save(fechamento_path)

    state["uploaded"] = True
    state["step1_done"] = False
    state["step2_done"] = False
    state["step3_done"] = False
    state["files"] = {
        "motoristas": motoristas_path,
        "fechamento": fechamento_path,
        "banco": "",
        "espelhos": "",
        "final": ""
    }
    save_state(state)

    flash("Arquivos enviados com sucesso. Agora gere o banco consolidado.", "ok")
    return redirect(url_for("index"))


@app.route("/step1", methods=["POST"])
def step1():
    if not is_logged_in():
        return redirect(url_for("login"))

    state = load_state()
    try:
        require_uploaded_files(state)
        ws = _ws_dir()

        banco_path = os.path.join(ws, "banco_consolidado.xlsx")
        gerar_banco_consolidado(
            motoristas_xlsx_path=state["files"]["motoristas"],
            fechamento_xlsx_path=state["files"]["fechamento"],
            saida_xlsx_path=banco_path,
        )

        state["files"]["banco"] = banco_path
        state["step1_done"] = True
        state["step2_done"] = False
        state["step3_done"] = False
        save_state(state)

        flash("Banco consolidado gerado. Agora gere os espelhos.", "ok")
    except Exception as e:
        flash(f"Erro no passo 1: {e}", "error")

    return redirect(url_for("index"))


@app.route("/step2", methods=["POST"])
def step2():
    if not is_logged_in():
        return redirect(url_for("login"))

    state = load_state()
    try:
        if not state.get("step1_done"):
            raise ValueError("Faça o passo 1 antes.")
        ws = _ws_dir()

        banco_path = state["files"].get("banco")
        if not banco_path or not os.path.exists(banco_path):
            raise FileNotFoundError("banco_consolidado.xlsx não encontrado.")

        espelhos_path = os.path.join(ws, "Espelhos_Motoristas.xlsx")
        gerar_espelhos_motoristas(
            banco_consolidado_xlsx_path=banco_path,
            modelo_xlsx_path=MODELO_PATH,
            saida_espelhos_xlsx_path=espelhos_path,
        )

        state["files"]["espelhos"] = espelhos_path
        state["step2_done"] = True
        state["step3_done"] = False
        save_state(state)

        flash("Espelhos gerados. Agora gere RESUMO e RESUMO TOTAL.", "ok")
    except Exception as e:
        flash(f"Erro no passo 2: {e}", "error")

    return redirect(url_for("index"))


@app.route("/step3", methods=["POST"])
def step3():
    if not is_logged_in():
        return redirect(url_for("login"))

    state = load_state()
    try:
        if not state.get("step2_done"):
            raise ValueError("Faça o passo 2 antes.")
        ws = _ws_dir()
        dl = _dl_dir()

        espelhos_path = state["files"].get("espelhos")
        banco_path = state["files"].get("banco")

        if not espelhos_path or not os.path.exists(espelhos_path):
            raise FileNotFoundError("Espelhos_Motoristas.xlsx não encontrado.")
        if not banco_path or not os.path.exists(banco_path):
            raise FileNotFoundError("banco_consolidado.xlsx não encontrado.")

        gerar_resumos(espelhos_xlsx_path=espelhos_path, banco_consolidado_xlsx_path=banco_path)

        final_path = os.path.join(dl, "Espelhos_Motoristas_FINAL.xlsx")
        if os.path.exists(final_path):
            os.remove(final_path)

        with open(espelhos_path, "rb") as src, open(final_path, "wb") as dst:
            dst.write(src.read())

        state["files"]["final"] = final_path
        state["step3_done"] = True
        save_state(state)

        flash("RESUMO e RESUMO TOTAL gerados. Download liberado.", "ok")
    except Exception as e:
        flash(f"Erro no passo 3: {e}", "error")

    return redirect(url_for("index"))


@app.route("/download", methods=["GET"])
def download():
    if not is_logged_in():
        return redirect(url_for("login"))

    state = load_state()
    final_path = state.get("files", {}).get("final")

    if not state.get("step3_done") or not final_path or not os.path.exists(final_path):
        flash("Arquivo final ainda não está pronto.", "error")
        return redirect(url_for("index"))

    return send_file(final_path, as_attachment=True, download_name="Espelhos_Motoristas.xlsx")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
