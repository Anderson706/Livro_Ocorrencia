import os
import base64
from io import BytesIO
from html import escape
from datetime import datetime, date
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, session, send_file
)
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func, case
from werkzeug.security import generate_password_hash, check_password_hash

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    Image, SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
)


from io import BytesIO
from datetime import datetime
import os

from flask import send_file
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image,
    KeepTogether
)


app = Flask(__name__)
app.config["SECRET_KEY"] = "dev-secret-change-me"

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, "livro_ocorrencias.db")
app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{DB_PATH}"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)


# =========================
# CONSTANTES / PADRÕES
# =========================
SITES_VALIDOS = {"PG", "SHEIN", "ADIDAS", "MELI", "SANGOBAN"}
TURNOS_VALIDOS = {"TURNO A", "TURNO B", "TURNO C", "ADM"}
TIPOS_VALIDOS = {
    "ROTINA",
    "DESVIO DE PROCESSO",
    "EXTRAVIO",
    "INCIDENTE",
    "MANUTENCAO",
    "MANUTENÇÃO",
    "PENDENCIA",
    "PENDÊNCIA",
    "INFORMATIVO",
}
PRIORIDADES_VALIDAS = {"BAIXA", "MEDIA", "MÉDIA", "ALTA", "CRITICA", "CRÍTICA"}
STATUS_VALIDOS = {"EM ABERTO", "EM ACOMPANHAMENTO", "FINALIZADO"}


# =========================
# MODELOS
# =========================
class User(db.Model):
    __tablename__ = "users"

    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(120), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False, default="USER")
    site = db.Column(db.String(80), nullable=True)
    is_active = db.Column(db.Boolean, nullable=False, default=True)
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.now)

    def set_password(self, senha: str):
        self.password_hash = generate_password_hash(senha)

    def check_password(self, senha: str) -> bool:
        return check_password_hash(self.password_hash, senha)


class OcorrenciaTurno(db.Model):
    __tablename__ = "ocorrencias_turno"

    id = db.Column(db.Integer, primary_key=True)
    data_ocorrencia = db.Column(db.Date, nullable=False, default=date.today)
    data_hora_registro = db.Column(db.DateTime, nullable=False, default=datetime.now)

    site = db.Column(db.String(80), nullable=False)
    turno = db.Column(db.String(30), nullable=False)
    setor = db.Column(db.String(100), nullable=False)
    tipo_ocorrencia = db.Column(db.String(100), nullable=False)
    prioridade = db.Column(db.String(20), nullable=False)

    responsavel_saida = db.Column(db.String(120), nullable=False)
    responsavel_entrada = db.Column(db.String(120), nullable=False)

    descricao = db.Column(db.Text, nullable=False)
    efetivo = db.Column(db.Text, nullable=False)
    assinatura = db.Column(db.Text, nullable=True)
    acoes_tomadas = db.Column(db.Text, nullable=True)
    pendencias = db.Column(db.Text, nullable=True)

    status = db.Column(db.String(40), nullable=False)

    criado_por = db.Column(db.String(120), nullable=True)
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.now)
    updated_at = db.Column(db.DateTime, nullable=True)

    def to_dict(self):
        return {
            "id": self.id,
            "data_ocorrencia": self.data_ocorrencia.strftime("%d/%m/%Y") if self.data_ocorrencia else "",
            "data_hora_registro": self.data_hora_registro.strftime("%d/%m/%Y %H:%M") if self.data_hora_registro else "",
            "site": self.site,
            "turno": self.turno,
            "setor": self.setor,
            "tipo_ocorrencia": self.tipo_ocorrencia,
            "prioridade": self.prioridade,
            "responsavel_saida": self.responsavel_saida,
            "responsavel_entrada": self.responsavel_entrada,
            "descricao": self.descricao,
            "efetivo": self.efetivo or "",
            "assinatura": self.assinatura or "",
            "acoes_tomadas": self.acoes_tomadas or "",
            "pendencias": self.pendencias or "",
            "status": self.status,
            "criado_por": self.criado_por or "",
            "created_at": self.created_at.strftime("%d/%m/%Y %H:%M") if self.created_at else "",
            "updated_at": self.updated_at.strftime("%d/%m/%Y %H:%M") if self.updated_at else "",
        }


def garantir_colunas_ocorrencias():
    with db.engine.connect() as conn:
        result = conn.execute(db.text("PRAGMA table_info(ocorrencias_turno)"))
        colunas = {row[1] for row in result.fetchall()}

        if "efetivo" not in colunas:
            conn.execute(db.text("ALTER TABLE ocorrencias_turno ADD COLUMN efetivo TEXT"))

        if "assinatura" not in colunas:
            conn.execute(db.text("ALTER TABLE ocorrencias_turno ADD COLUMN assinatura TEXT"))

        conn.commit()



def garantir_colunas_ocorrencias():
    with db.engine.connect() as conn:
        result = conn.execute(db.text("PRAGMA table_info(ocorrencias_turno)"))
        colunas = {row[1] for row in result.fetchall()}

        if "efetivo" not in colunas:
            conn.execute(db.text("ALTER TABLE ocorrencias_turno ADD COLUMN efetivo TEXT"))

        if "assinatura" not in colunas:
            conn.execute(db.text("ALTER TABLE ocorrencias_turno ADD COLUMN assinatura VARCHAR(120)"))

        conn.commit()

# =========================
# DECORATORS
# =========================
def login_required(funcao):
    @wraps(funcao)
    def wrapper(*args, **kwargs):
        if not session.get("user_id"):
            flash("Faça login para acessar o sistema.", "warning")
            return redirect(url_for("login"))
        return funcao(*args, **kwargs)
    return wrapper


def admin_required(funcao):
    @wraps(funcao)
    def wrapper(*args, **kwargs):
        if not session.get("user_id"):
            flash("Faça login para acessar o sistema.", "warning")
            return redirect(url_for("login"))
        if session.get("user_role") != "ADMIN":
            flash("Apenas administradores podem acessar esta funcionalidade.", "danger")
            return redirect(url_for("index"))
        return funcao(*args, **kwargs)
    return wrapper


# =========================
# HELPERS
# =========================
def normalizar_texto(valor: str) -> str:
    return (valor or "").strip().upper()


def normalizar_prioridade(valor: str) -> str:
    valor = normalizar_texto(valor)
    return "CRITICA" if valor == "CRÍTICA" else "MEDIA" if valor == "MÉDIA" else valor


def normalizar_tipo(valor: str) -> str:
    valor = normalizar_texto(valor)
    if valor == "MANUTENÇÃO":
        return "MANUTENCAO"
    if valor == "PENDÊNCIA":
        return "PENDENCIA"
    return valor


def parse_date_or_none(value: str):
    value = (value or "").strip()
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError:
        return None


def parse_datetime_local(value: str):
    value = (value or "").strip()
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%dT%H:%M")
    except ValueError:
        return None


def format_date_input(value):
    if not value:
        return ""
    if isinstance(value, str):
        return value
    return value.strftime("%Y-%m-%d")


def format_datetime_local_input(value):
    if not value:
        return ""
    if isinstance(value, str):
        return value
    return value.strftime("%Y-%m-%dT%H:%M")


def pdf_safe(texto):
    return escape(texto or "-")


def get_filtros_ocorrencias():
    data_inicial = (request.args.get("data_inicial") or "").strip()
    data_final = (request.args.get("data_final") or "").strip()
    turno = normalizar_texto(request.args.get("turno"))
    status = normalizar_texto(request.args.get("status"))
    site = normalizar_texto(request.args.get("site"))

    query = OcorrenciaTurno.query

    if data_inicial:
        di = parse_date_or_none(data_inicial)
        if di:
            query = query.filter(OcorrenciaTurno.data_ocorrencia >= di)

    if data_final:
        df = parse_date_or_none(data_final)
        if df:
            query = query.filter(OcorrenciaTurno.data_ocorrencia <= df)

    if turno:
        query = query.filter(OcorrenciaTurno.turno == turno)

    if status:
        query = query.filter(OcorrenciaTurno.status == status)

    if site:
        query = query.filter(OcorrenciaTurno.site == site)

    query = query.order_by(
        OcorrenciaTurno.data_hora_registro.desc(),
        OcorrenciaTurno.id.desc()
    )

    filtros = {
        "data_inicial": data_inicial,
        "data_final": data_final,
        "turno": turno,
        "status": status,
        "site": site,
    }
    return query, filtros


def resumo_cards():
    hoje = date.today()

    ocorrencias_dia = (
        db.session.query(func.count(OcorrenciaTurno.id))
        .filter(OcorrenciaTurno.data_ocorrencia == hoje)
        .scalar()
    ) or 0

    pendencias_abertas = (
        db.session.query(func.count(OcorrenciaTurno.id))
        .filter(OcorrenciaTurno.status.in_(["EM ABERTO", "EM ACOMPANHAMENTO"]))
        .scalar()
    ) or 0

    turnos_registrados = (
        db.session.query(func.count(func.distinct(OcorrenciaTurno.turno)))
        .filter(OcorrenciaTurno.data_ocorrencia == hoje)
        .scalar()
    ) or 0

    ocorrencias_criticas = (
        db.session.query(func.count(OcorrenciaTurno.id))
        .filter(OcorrenciaTurno.prioridade == "CRITICA")
        .scalar()
    ) or 0

    return {
        "ocorrencias_dia": ocorrencias_dia,
        "pendencias_abertas": pendencias_abertas,
        "turnos_registrados": turnos_registrados,
        "ocorrencias_criticas": ocorrencias_criticas,
    }


def pode_criar_admin_publicamente():
    admin_existe = User.query.filter_by(role="ADMIN").first()
    return admin_existe is None


def ordenar_turnos_query():
    return case(
        (OcorrenciaTurno.turno == "TURNO A", 1),
        (OcorrenciaTurno.turno == "TURNO B", 2),
        (OcorrenciaTurno.turno == "TURNO C", 3),
        (OcorrenciaTurno.turno == "ADM", 4),
        else_=99
    )


def ordenar_prioridade_query():
    return case(
        (OcorrenciaTurno.prioridade == "CRITICA", 1),
        (OcorrenciaTurno.prioridade == "ALTA", 2),
        (OcorrenciaTurno.prioridade == "MEDIA", 3),
        (OcorrenciaTurno.prioridade == "BAIXA", 4),
        else_=99
    )


# =========================
# AUTH
# =========================
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = (request.form.get("email") or "").strip().lower()
        senha = request.form.get("password") or ""

        user = User.query.filter_by(email=email, is_active=True).first()

        if not user or not user.check_password(senha):
            flash("E-mail ou senha inválidos.", "danger")
            return redirect(url_for("login"))

        session["user_id"] = user.id
        session["username"] = user.nome
        session["user_role"] = user.role
        session["user_site"] = user.site or ""

        flash("Login realizado com sucesso.", "success")
        return redirect(url_for("index"))

    return render_template(
        "login.html",
        permitir_admin_publico=pode_criar_admin_publicamente()
    )


@app.route("/logout")
def logout():
    session.clear()
    flash("Sessão encerrada com sucesso.", "success")
    return redirect(url_for("login"))


# =========================
# CRIAÇÃO DE USUÁRIO
# =========================
@app.route("/criar-usuario", methods=["GET", "POST"])
def criar_usuario():
    admin_publico_liberado = pode_criar_admin_publicamente()
    usuario_logado_admin = session.get("user_role") == "ADMIN"

    if request.method == "POST":
        nome = (request.form.get("nome") or "").strip()
        email = (request.form.get("email") or "").strip().lower()
        senha = request.form.get("senha") or ""
        role_solicitado = normalizar_texto(request.form.get("role") or "USER")
        site = normalizar_texto(request.form.get("site"))

        if not nome or not email or not senha:
            flash("Preencha nome, e-mail e senha.", "danger")
            return redirect(url_for("criar_usuario"))

        if role_solicitado not in {"ADMIN", "USER"}:
            role_solicitado = "USER"

        if role_solicitado == "ADMIN" and not (admin_publico_liberado or usuario_logado_admin):
            flash("Somente administradores podem criar outro administrador.", "danger")
            return redirect(url_for("criar_usuario"))

        existe = User.query.filter_by(email=email).first()
        if existe:
            flash("Já existe um usuário com esse e-mail.", "warning")
            return redirect(url_for("criar_usuario"))

        novo = User(
            nome=nome,
            email=email,
            role=role_solicitado,
            site=site,
            is_active=True
        )
        novo.set_password(senha)

        db.session.add(novo)
        db.session.commit()

        flash("Usuário criado com sucesso.", "success")

        if not session.get("user_id"):
            return redirect(url_for("login"))

        return redirect(url_for("usuarios"))

    return render_template(
        "criar_usuario.html",
        permitir_admin_publico=admin_publico_liberado,
        usuario_logado_admin=usuario_logado_admin
    )


# =========================
# USUÁRIOS
# =========================
@app.route("/usuarios")
@admin_required
def usuarios():
    lista = User.query.order_by(User.nome.asc()).all()
    return render_template("usuarios.html", usuarios=lista)


@app.route("/usuarios/novo", methods=["GET", "POST"])
@admin_required
def novo_usuario():
    if request.method == "POST":
        nome = (request.form.get("nome") or "").strip()
        email = (request.form.get("email") or "").strip().lower()
        senha = request.form.get("senha") or ""
        role = normalizar_texto(request.form.get("role") or "USER")
        site = normalizar_texto(request.form.get("site"))

        if not nome or not email or not senha:
            flash("Preencha nome, e-mail e senha.", "danger")
            return redirect(url_for("novo_usuario"))

        if role not in {"ADMIN", "USER"}:
            role = "USER"

        existe = User.query.filter_by(email=email).first()
        if existe:
            flash("Já existe um usuário com esse e-mail.", "warning")
            return redirect(url_for("novo_usuario"))

        user = User(
            nome=nome,
            email=email,
            role=role,
            site=site,
            is_active=True,
        )
        user.set_password(senha)

        db.session.add(user)
        db.session.commit()

        flash("Usuário criado com sucesso.", "success")
        return redirect(url_for("usuarios"))

    return render_template("usuario_form.html", usuario=None)


@app.route("/usuarios/<int:user_id>/editar", methods=["GET", "POST"])
@admin_required
def editar_usuario(user_id):
    usuario = User.query.get_or_404(user_id)

    if request.method == "POST":
        nome = (request.form.get("nome") or "").strip()
        email = (request.form.get("email") or "").strip().lower()
        senha = (request.form.get("senha") or "").strip()
        role = normalizar_texto(request.form.get("role") or "USER")
        site = normalizar_texto(request.form.get("site"))
        is_active = (request.form.get("is_active") or "").strip() == "1"

        if not nome or not email:
            flash("Preencha nome e e-mail.", "danger")
            return redirect(url_for("editar_usuario", user_id=user_id))

        if role not in {"ADMIN", "USER"}:
            role = "USER"

        existe = User.query.filter(User.email == email, User.id != usuario.id).first()
        if existe:
            flash("Já existe outro usuário com esse e-mail.", "warning")
            return redirect(url_for("editar_usuario", user_id=user_id))

        usuario.nome = nome
        usuario.email = email
        usuario.role = role
        usuario.site = site
        usuario.is_active = is_active

        if senha:
            usuario.set_password(senha)

        db.session.commit()

        if usuario.id == session.get("user_id"):
            session["username"] = usuario.nome
            session["user_role"] = usuario.role
            session["user_site"] = usuario.site or ""

        flash("Usuário atualizado com sucesso.", "success")
        return redirect(url_for("usuarios"))

    return render_template("usuario_form.html", usuario=usuario)


@app.route("/usuarios/<int:user_id>/excluir", methods=["POST"])
@admin_required
def excluir_usuario(user_id):
    usuario = User.query.get_or_404(user_id)

    if usuario.id == session.get("user_id"):
        flash("Você não pode excluir o próprio usuário logado.", "warning")
        return redirect(url_for("usuarios"))

    db.session.delete(usuario)
    db.session.commit()
    flash("Usuário excluído com sucesso.", "success")
    return redirect(url_for("usuarios"))


# =========================
# OCORRÊNCIAS
# =========================
@app.route("/", methods=["GET"])
@login_required
def index():
    query, filtros = get_filtros_ocorrencias()
    ocorrencias_db = query.all()
    ocorrencias = [o.to_dict() for o in ocorrencias_db]
    ultima_ocorrencia = ocorrencias[0] if ocorrencias else None

    ultimo_id = db.session.query(func.max(OcorrenciaTurno.id)).scalar() or 0
    proximo_id_previsto = ultimo_id + 1

    return render_template(
        "livro_ocorrencia.html",
        resumo=resumo_cards(),
        ultima_ocorrencia=ultima_ocorrencia,
        ocorrencias=ocorrencias,
        filtros=filtros,
        hoje=date.today().strftime("%Y-%m-%d"),
        proximo_id_previsto=proximo_id_previsto
    )


@app.route("/salvar-ocorrencia-turno", methods=["POST"])
@login_required
def salvar_ocorrencia_turno():
    try:
        data_ocorrencia = parse_date_or_none(request.form.get("data_ocorrencia"))
        data_hora_registro = parse_datetime_local(request.form.get("data_hora_registro"))

        site = normalizar_texto(request.form.get("site"))
        turno = normalizar_texto(request.form.get("turno"))
        setor = normalizar_texto(request.form.get("setor"))
        tipo_ocorrencia = normalizar_tipo(request.form.get("tipo_ocorrencia"))
        prioridade = normalizar_prioridade(request.form.get("prioridade"))
        responsavel_saida = (request.form.get("responsavel_saida") or "").strip()
        responsavel_entrada = (request.form.get("responsavel_entrada") or "").strip()
        descricao = (request.form.get("descricao") or "").strip()
        efetivo = (request.form.get("efetivo") or "").strip()
        assinatura = request.form.get("assinatura") or ""
        acoes_tomadas = (request.form.get("acoes_tomadas") or "").strip()
        pendencias = (request.form.get("pendencias") or "").strip()
        status = normalizar_texto(request.form.get("status"))

        if not all([
            data_ocorrencia, data_hora_registro, site, turno, setor,
            tipo_ocorrencia, prioridade, responsavel_saida,
            responsavel_entrada, descricao, efetivo, status
        ]):
            flash("Preencha todos os campos obrigatórios.", "danger")
            return redirect(url_for("index"))

        if site not in SITES_VALIDOS:
            flash("Site inválido.", "danger")
            return redirect(url_for("index"))

        if turno not in TURNOS_VALIDOS:
            flash("Turno inválido.", "danger")
            return redirect(url_for("index"))

        if prioridade not in {"BAIXA", "MEDIA", "ALTA", "CRITICA"}:
            flash("Prioridade inválida.", "danger")
            return redirect(url_for("index"))

        if status not in STATUS_VALIDOS:
            flash("Status inválido.", "danger")
            return redirect(url_for("index"))

        nova = OcorrenciaTurno(
            data_ocorrencia=data_ocorrencia,
            data_hora_registro=data_hora_registro,
            site=site,
            turno=turno,
            setor=setor,
            tipo_ocorrencia=tipo_ocorrencia,
            prioridade=prioridade,
            responsavel_saida=responsavel_saida,
            responsavel_entrada=responsavel_entrada,
            descricao=descricao,
            efetivo=efetivo,
            assinatura=assinatura or None,
            acoes_tomadas=acoes_tomadas or None,
            pendencias=pendencias or None,
            status=status,
            criado_por=session.get("username", "Usuário")
        )

        db.session.add(nova)
        db.session.commit()

        flash("Ocorrência registrada com sucesso.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Erro ao salvar ocorrência: {e}", "danger")

    return redirect(url_for("index"))


@app.route("/ocorrencias/<int:ocorrencia_id>/editar", methods=["GET", "POST"])
@login_required
def editar_ocorrencia(ocorrencia_id):
    ocorrencia = OcorrenciaTurno.query.get_or_404(ocorrencia_id)

    if request.method == "POST":
        try:
            ocorrencia.data_ocorrencia = parse_date_or_none(request.form.get("data_ocorrencia"))
            ocorrencia.data_hora_registro = parse_datetime_local(request.form.get("data_hora_registro"))

            ocorrencia.site = normalizar_texto(request.form.get("site"))
            ocorrencia.turno = normalizar_texto(request.form.get("turno"))
            ocorrencia.setor = normalizar_texto(request.form.get("setor"))
            ocorrencia.tipo_ocorrencia = normalizar_tipo(request.form.get("tipo_ocorrencia"))
            ocorrencia.prioridade = normalizar_prioridade(request.form.get("prioridade"))
            ocorrencia.responsavel_saida = (request.form.get("responsavel_saida") or "").strip()
            ocorrencia.responsavel_entrada = (request.form.get("responsavel_entrada") or "").strip()
            ocorrencia.descricao = (request.form.get("descricao") or "").strip()
            ocorrencia.efetivo = (request.form.get("efetivo") or "").strip()

            assinatura_recebida = request.form.get("assinatura") or ""
            if assinatura_recebida:
                ocorrencia.assinatura = assinatura_recebida

            ocorrencia.acoes_tomadas = (request.form.get("acoes_tomadas") or "").strip() or None
            ocorrencia.pendencias = (request.form.get("pendencias") or "").strip() or None
            ocorrencia.status = normalizar_texto(request.form.get("status"))
            ocorrencia.updated_at = datetime.now()

            if not all([
                ocorrencia.data_ocorrencia, ocorrencia.data_hora_registro, ocorrencia.site,
                ocorrencia.turno, ocorrencia.setor, ocorrencia.tipo_ocorrencia,
                ocorrencia.prioridade, ocorrencia.responsavel_saida,
                ocorrencia.responsavel_entrada, ocorrencia.descricao,
                ocorrencia.efetivo, ocorrencia.status
            ]):
                flash("Preencha todos os campos obrigatórios.", "danger")
                return redirect(url_for("editar_ocorrencia", ocorrencia_id=ocorrencia.id))

            if ocorrencia.site not in SITES_VALIDOS:
                flash("Site inválido.", "danger")
                return redirect(url_for("editar_ocorrencia", ocorrencia_id=ocorrencia.id))

            if ocorrencia.turno not in TURNOS_VALIDOS:
                flash("Turno inválido.", "danger")
                return redirect(url_for("editar_ocorrencia", ocorrencia_id=ocorrencia.id))

            if ocorrencia.prioridade not in {"BAIXA", "MEDIA", "ALTA", "CRITICA"}:
                flash("Prioridade inválida.", "danger")
                return redirect(url_for("editar_ocorrencia", ocorrencia_id=ocorrencia.id))

            if ocorrencia.status not in STATUS_VALIDOS:
                flash("Status inválido.", "danger")
                return redirect(url_for("editar_ocorrencia", ocorrencia_id=ocorrencia.id))

            db.session.commit()
            flash("Ocorrência atualizada com sucesso.", "success")
            return redirect(url_for("index"))

        except Exception as e:
            db.session.rollback()
            flash(f"Erro ao atualizar ocorrência: {e}", "danger")
            return redirect(url_for("editar_ocorrencia", ocorrencia_id=ocorrencia.id))

    return render_template(
        "ocorrencia_form.html",
        ocorrencia=ocorrencia,
        data_ocorrencia_value=format_date_input(ocorrencia.data_ocorrencia),
        data_hora_value=format_datetime_local_input(ocorrencia.data_hora_registro)
    )

@app.route("/ocorrencias/<int:ocorrencia_id>/fechar", methods=["POST"])
@login_required
def fechar_ocorrencia(ocorrencia_id):
    ocorrencia = OcorrenciaTurno.query.get_or_404(ocorrencia_id)

    try:
        ocorrencia.status = "FINALIZADO"
        ocorrencia.updated_at = datetime.now()
        db.session.commit()
        flash("Ocorrência finalizada com sucesso.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Erro ao finalizar ocorrência: {e}", "danger")

    return redirect(url_for("index"))


@app.route("/ocorrencias/<int:ocorrencia_id>/excluir", methods=["POST"])
@login_required
def excluir_ocorrencia(ocorrencia_id):
    ocorrencia = OcorrenciaTurno.query.get_or_404(ocorrencia_id)
    try:
        db.session.delete(ocorrencia)
        db.session.commit()
        flash("Ocorrência excluída com sucesso.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Erro ao excluir ocorrência: {e}", "danger")
    return redirect(url_for("index"))


# =========================
# DASHBOARD
# =========================
@app.route("/dashboard-ocorrencias")
@login_required

def dashboard_ocorrencias():

    total = db.session.query(func.count(OcorrenciaTurno.id)).scalar() or 0
    em_aberto = db.session.query(func.count(OcorrenciaTurno.id)).filter(
        OcorrenciaTurno.status == "EM ABERTO"
    ).scalar() or 0
    acompanhamento = db.session.query(func.count(OcorrenciaTurno.id)).filter(
        OcorrenciaTurno.status == "EM ACOMPANHAMENTO"
    ).scalar() or 0
    finalizado = db.session.query(func.count(OcorrenciaTurno.id)).filter(
        OcorrenciaTurno.status == "FINALIZADO"
    ).scalar() or 0

    por_turno = (
        db.session.query(OcorrenciaTurno.turno, func.count(OcorrenciaTurno.id))
        .group_by(OcorrenciaTurno.turno)
        .order_by(ordenar_turnos_query())
        .all()
    )

    por_prioridade = (
        db.session.query(OcorrenciaTurno.prioridade, func.count(OcorrenciaTurno.id))
        .group_by(OcorrenciaTurno.prioridade)
        .order_by(ordenar_prioridade_query())
        .all()
    )

    por_site = (
        db.session.query(OcorrenciaTurno.site, func.count(OcorrenciaTurno.id))
        .group_by(OcorrenciaTurno.site)
        .order_by(OcorrenciaTurno.site.asc())
        .all()
    )

    return render_template(
        "dashboard_ocorrencias.html",
        total=total,
        em_aberto=em_aberto,
        acompanhamento=acompanhamento,
        finalizado=finalizado,
        por_turno=por_turno,
        por_prioridade=por_prioridade,
        por_site=por_site,
    )


# =========================
# EXPORTAÇÃO EXCEL
# =========================
@app.route("/export-ocorrencias-excel")
@login_required
def export_ocorrencias_excel():
    query, _ = get_filtros_ocorrencias()
    rows = query.all()

    wb = Workbook()
    ws = wb.active
    ws.title = "Livro de Ocorrências"

    headers = [
        "ID", "Data da Ocorrência", "Data/Hora Registro", "Site", "Turno", "Setor",
        "Tipo de Ocorrência", "Prioridade", "Responsável Saída", "Responsável Entrada",
        "Efetivo", "Descrição", "Ações Tomadas", "Pendências", "Assinatura", "Status",
        "Criado por", "Criado em", "Atualizado em"
    ]
    ws.append(headers)

    fill_header = PatternFill("solid", fgColor="FFCC00")
    font_header = Font(bold=True, color="000000")

    for col_num, _ in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r in rows:
        ws.append([
            r.id,
            r.data_ocorrencia.strftime("%d/%m/%Y") if r.data_ocorrencia else "",
            r.data_hora_registro.strftime("%d/%m/%Y %H:%M") if r.data_hora_registro else "",
            r.site,
            r.turno,
            r.setor,
            r.tipo_ocorrencia,
            r.prioridade,
            r.responsavel_saida,
            r.responsavel_entrada,
            r.efetivo or "",
            r.descricao,
            r.acoes_tomadas or "",
            r.pendencias or "",
            r.criado_por or "",
            r.status,
            r.criado_por or "",
            r.created_at.strftime("%d/%m/%Y %H:%M") if r.created_at else "",
            r.updated_at.strftime("%d/%m/%Y %H:%M") if r.updated_at else "",
        ])

    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            value = str(cell.value) if cell.value is not None else ""
            max_length = max(max_length, len(value))
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        ws.column_dimensions[col_letter].width = min(max_length + 2, 40)

    ws.freeze_panes = "A2"

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    nome_arquivo = f"livro_ocorrencias_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=nome_arquivo,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def assinatura_base64_para_image(assinatura_b64, largura_mm=60, altura_mm=22):
    if not assinatura_b64:
        return None

    try:
        if "," in assinatura_b64:
            _, encoded = assinatura_b64.split(",", 1)
        else:
            encoded = assinatura_b64

        image_bytes = base64.b64decode(encoded)
        buffer = BytesIO(image_bytes)
        return Image(buffer, width=largura_mm * mm, height=altura_mm * mm)
    except Exception:
        return None



@app.route("/ocorrencias/<int:ocorrencia_id>/pdf")
@login_required
def export_ocorrencia_individual_pdf(ocorrencia_id):
    ocorrencia = OcorrenciaTurno.query.get_or_404(ocorrencia_id)

    output = BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=A4,
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=24 * mm,
        bottomMargin=16 * mm
    )

    styles = getSampleStyleSheet()

    # =========================
    # HELPERS
    # =========================
    def v(valor, default="-"):
        if valor is None:
            return default
        valor = str(valor).strip()
        return pdf_safe(valor) if valor else default

    def dt(valor, fmt="%d/%m/%Y %H:%M"):
        return valor.strftime(fmt) if valor else "-"

    def prioridade_cfg(prioridade):
        p = (prioridade or "").strip().lower()
        if p in ["alta", "critica", "crítica"]:
            return {
                "bg": colors.HexColor("#FDECEC"),
                "fg": colors.HexColor("#B42318"),
                "label": v(prioridade)
            }
        elif p in ["média", "media"]:
            return {
                "bg": colors.HexColor("#FFF4DB"),
                "fg": colors.HexColor("#9A6700"),
                "label": v(prioridade)
            }
        elif p in ["baixa"]:
            return {
                "bg": colors.HexColor("#EAF4FF"),
                "fg": colors.HexColor("#175CD3"),
                "label": v(prioridade)
            }
        return {
            "bg": colors.HexColor("#F3F4F6"),
            "fg": colors.HexColor("#374151"),
            "label": v(prioridade)
        }

    def status_cfg(status):
        s = (status or "").strip().lower()
        if s in ["concluída", "concluida", "fechada", "finalizada"]:
            return {
                "bg": colors.HexColor("#E8F7EE"),
                "fg": colors.HexColor("#146C43"),
                "label": v(status)
            }
        elif s in ["aberta", "pendente", "em andamento"]:
            return {
                "bg": colors.HexColor("#FFF4DB"),
                "fg": colors.HexColor("#9A6700"),
                "label": v(status)
            }
        elif s in ["crítica", "critica", "atrasada"]:
            return {
                "bg": colors.HexColor("#FDE8EA"),
                "fg": colors.HexColor("#B42318"),
                "label": v(status)
            }
        return {
            "bg": colors.HexColor("#EEF2F6"),
            "fg": colors.HexColor("#344054"),
            "label": v(status)
        }

    # =========================
    # ESTILOS
    # =========================
    title_style = ParagraphStyle(
        "TituloDHL",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=20,
        leading=24,
        textColor=colors.HexColor("#D40511"),
        alignment=TA_LEFT,
        spaceAfter=4,
    )

    subtitle_style = ParagraphStyle(
        "SubTituloDHL",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#667085"),
        alignment=TA_LEFT,
        spaceAfter=8,
    )

    section_title_style = ParagraphStyle(
        "SectionTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10,
        leading=12,
        textColor=colors.white,
        alignment=TA_LEFT,
    )

    label_style = ParagraphStyle(
        "Label",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=11,
        textColor=colors.HexColor("#111827"),
        spaceAfter=1,
    )

    value_style = ParagraphStyle(
        "Value",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=11,
        textColor=colors.HexColor("#344054"),
    )

    block_title_style = ParagraphStyle(
        "BlockTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=13,
        textColor=colors.HexColor("#111827"),
        spaceAfter=6,
    )

    text_style = ParagraphStyle(
        "Texto",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10,
        leading=15,
        textColor=colors.HexColor("#333333"),
        spaceAfter=0,
    )

    badge_center_style = ParagraphStyle(
        "BadgeCenter",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10,
        leading=12,
        alignment=TA_CENTER,
    )

    kpi_number_style = ParagraphStyle(
        "KPINumber",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=20,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#111111"),
    )

    kpi_label_style = ParagraphStyle(
        "KPILabel",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#6B7280"),
    )

    footer_style = ParagraphStyle(
        "FooterStyle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        textColor=colors.HexColor("#6B7280"),
        alignment=TA_RIGHT,
    )

    # =========================
    # HEADER / FOOTER
    # =========================
    def draw_header_footer(canvas, doc):
        canvas.saveState()
        width, height = A4

        # Barra superior vermelha
        canvas.setFillColor(colors.HexColor("#D40511"))
        canvas.rect(0, height - 14 * mm, width, 14 * mm, fill=1, stroke=0)

        # Faixa amarela
        canvas.setFillColor(colors.HexColor("#FFCC00"))
        canvas.rect(0, height - 16 * mm, width, 2 * mm, fill=1, stroke=0)

        # Texto cabeçalho
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica-Bold", 11)
        canvas.drawString(14 * mm, height - 9 * mm, "DHL SECURITY")

        canvas.setFont("Helvetica", 8)
        canvas.drawRightString(
            width - 14 * mm,
            height - 9 * mm,
            f"Documento emitido em {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        )

        # Rodapé
        canvas.setStrokeColor(colors.HexColor("#D0D5DD"))
        canvas.setLineWidth(0.4)
        canvas.line(14 * mm, 11 * mm, width - 14 * mm, 11 * mm)

        canvas.setFillColor(colors.HexColor("#667085"))
        canvas.setFont("Helvetica", 8)
        canvas.drawString(14 * mm, 7 * mm, "Livro de Ocorrência • Relatório Individual")
        canvas.drawRightString(width - 14 * mm, 7 * mm, f"Página {canvas.getPageNumber()}")

        canvas.restoreState()

    # =========================
    # ELEMENTOS
    # =========================
    elementos = []

    # Logo
    # Troque o nome do arquivo abaixo se seu logotipo estiver com outro nome.
    possiveis_logos = [
        os.path.join(app.static_folder, "logo.png"),
        os.path.join(app.static_folder, "logo.jpg"),
        os.path.join(app.static_folder, "logo.jpeg"),
        os.path.join(app.static_folder, "logo_dhl.png"),
        os.path.join(app.static_folder, "leroy_secu.png"),
    ]

    logo_path = next((p for p in possiveis_logos if os.path.exists(p)), None)
    if logo_path:
        try:
            logo = Image(logo_path, width=42 * mm, height=16 * mm)
            elementos.append(logo)
            elementos.append(Spacer(1, 6))
        except Exception:
            pass

    # Título
    elementos.append(Paragraph("RELATÓRIO INDIVIDUAL DE OCORRÊNCIA", title_style))
    elementos.append(Paragraph(
        "Documento detalhado com as informações completas do registro operacional.",
        subtitle_style
    ))
    elementos.append(Spacer(1, 4))

    # =========================
    # KPIs SUPERIORES
    # =========================
    prio = prioridade_cfg(ocorrencia.prioridade)
    stat = status_cfg(ocorrencia.status)

    card_id = Table([
        [Paragraph(f"#{ocorrencia.id}", kpi_number_style)],
        [Paragraph("ID DA OCORRÊNCIA", kpi_label_style)]
    ], colWidths=[44 * mm])

    card_id.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FFF8DB")),
        ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#FFCC00")),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))

    card_status = Table([
        [Paragraph(f'<font color="{stat["fg"]}"><b>{stat["label"]}</b></font>', badge_center_style)],
        [Paragraph("STATUS", kpi_label_style)]
    ], colWidths=[52 * mm])

    card_status.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), stat["bg"]),
        ("BACKGROUND", (0, 1), (0, 1), colors.white),
        ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#D0D5DD")),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))

    card_prio = Table([
        [Paragraph(f'<font color="{prio["fg"]}"><b>{prio["label"]}</b></font>', badge_center_style)],
        [Paragraph("PRIORIDADE", kpi_label_style)]
    ], colWidths=[52 * mm])

    card_prio.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), prio["bg"]),
        ("BACKGROUND", (0, 1), (0, 1), colors.white),
        ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#D0D5DD")),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))

    resumo_topo = Table(
        [[card_id, card_status, card_prio]],
        colWidths=[48 * mm, 58 * mm, 58 * mm]
    )
    resumo_topo.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))

    elementos.append(KeepTogether(resumo_topo))
    elementos.append(Spacer(1, 10))

    # =========================
    # BLOCO DADOS GERAIS
    # =========================
    dados_principais = [
        [Paragraph("<b>Data da ocorrência</b>", label_style),
         Paragraph(v(dt(ocorrencia.data_ocorrencia, "%d/%m/%Y")), value_style)],
        [Paragraph("<b>Data/Hora do registro</b>", label_style),
         Paragraph(v(dt(ocorrencia.data_hora_registro)), value_style)],
        [Paragraph("<b>Site</b>", label_style), Paragraph(v(ocorrencia.site), value_style)],
        [Paragraph("<b>Turno</b>", label_style), Paragraph(v(ocorrencia.turno), value_style)],
        [Paragraph("<b>Setor</b>", label_style), Paragraph(v(ocorrencia.setor), value_style)],
        [Paragraph("<b>Tipo de ocorrência</b>", label_style), Paragraph(v(ocorrencia.tipo_ocorrencia), value_style)],
        [Paragraph("<b>Responsável saída</b>", label_style), Paragraph(v(ocorrencia.responsavel_saida), value_style)],
        [Paragraph("<b>Responsável entrada</b>", label_style),
         Paragraph(v(ocorrencia.responsavel_entrada), value_style)],
        [Paragraph("<b>Efetivo</b>", label_style), Paragraph(v(ocorrencia.efetivo), value_style)],
        [Paragraph("<b>Criado por</b>", label_style), Paragraph(v(ocorrencia.criado_por), value_style)],
        [Paragraph("<b>Criado em</b>", label_style), Paragraph(v(dt(ocorrencia.created_at)), value_style)],
        [Paragraph("<b>Atualizado em</b>", label_style), Paragraph(v(dt(ocorrencia.updated_at)), value_style)],
    ]

    cabecalho_info = Table(
        [[Paragraph("DADOS GERAIS DA OCORRÊNCIA", section_title_style)]],
        colWidths=[182 * mm]
    )
    cabecalho_info.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#D40511")),
        ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))

    tabela_info = Table(dados_principais, colWidths=[58 * mm, 124 * mm])
    tabela_info.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.white),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, colors.HexColor("#FAFAFA")]),
        ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#D0D5DD")),
        ("INNERGRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#EAECF0")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))

    elementos.append(KeepTogether(cabecalho_info))
    elementos.append(tabela_info)
    elementos.append(Spacer(1, 12))

    # =========================
    # SEÇÕES TEXTUAIS
    # =========================
    def bloco_texto(titulo, conteudo):
        cab = Table(
            [[Paragraph(titulo, section_title_style)]],
            colWidths=[182 * mm]
        )
        cab.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#111827")),
            ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))

        corpo = Table(
            [[Paragraph(v(conteudo), text_style)]],
            colWidths=[182 * mm]
        )
        corpo.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.white),
            ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#D0D5DD")),
            ("LEFTPADDING", (0, 0), (-1, -1), 10),
            ("RIGHTPADDING", (0, 0), (-1, -1), 10),
            ("TOPPADDING", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ]))

        return [KeepTogether(cab), corpo, Spacer(1, 10)]

    elementos.extend(bloco_texto("DESCRIÇÃO DA OCORRÊNCIA", ocorrencia.descricao))
    elementos.extend(bloco_texto("AÇÕES TOMADAS", ocorrencia.acoes_tomadas))
    elementos.extend(bloco_texto("PENDÊNCIAS", ocorrencia.pendencias))

    assinatura_img = assinatura_base64_para_image(ocorrencia.assinatura)
    if assinatura_img:
        assinatura_titulo = Table(
            [[Paragraph("ASSINATURA", section_title_style)]],
            colWidths=[182 * mm]
        )
        assinatura_titulo.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#111827")),
            ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))

        assinatura_corpo = Table(
            [[assinatura_img]],
            colWidths=[182 * mm]
        )
        assinatura_corpo.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.white),
            ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#D0D5DD")),
            ("LEFTPADDING", (0, 0), (-1, -1), 10),
            ("RIGHTPADDING", (0, 0), (-1, -1), 10),
            ("TOPPADDING", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ]))

        elementos.append(KeepTogether(assinatura_titulo))
        elementos.append(assinatura_corpo)
        elementos.append(Spacer(1, 10))

    # =========================
    # BLOCO FINAL
    # =========================
    emissao = Table(
        [[Paragraph(
            f"Documento gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')}",
            footer_style
        )]],
        colWidths=[182 * mm]
    )
    emissao.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F9FAFB")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#E5E7EB")),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    elementos.append(emissao)

    # =========================
    # BUILD
    # =========================
    doc.build(
        elementos,
        onFirstPage=draw_header_footer,
        onLaterPages=draw_header_footer
    )

    output.seek(0)
    nome_arquivo = f"ocorrencia_{ocorrencia.id}.pdf"

    return send_file(
        output,
        as_attachment=True,
        download_name=nome_arquivo,
        mimetype="application/pdf"
    )

# =========================
# EXPORTAÇÃO PDF GERAL
# =========================
from io import BytesIO
from datetime import datetime

from flask import send_file
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    KeepTogether
)

@app.route("/export-ocorrencias-pdf")
@login_required
def export_ocorrencias_pdf():
    query, filtros = get_filtros_ocorrencias()
    rows = query.all()

    output = BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(A4),
        leftMargin=10 * mm,
        rightMargin=10 * mm,
        topMargin=18 * mm,
        bottomMargin=14 * mm
    )

    styles = getSampleStyleSheet()

    # =========================
    # ESTILOS
    # =========================
    title_style = ParagraphStyle(
        "DHLTitle",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=19,
        leading=22,
        textColor=colors.HexColor("#D40511"),
        alignment=TA_LEFT,
        spaceAfter=4,
    )

    subtitle_style = ParagraphStyle(
        "DHLSubTitle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#444444"),
        alignment=TA_LEFT,
        spaceAfter=8,
    )

    section_label_style = ParagraphStyle(
        "SectionLabel",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=11,
        textColor=colors.white,
        alignment=TA_LEFT,
    )

    info_style = ParagraphStyle(
        "InfoStyle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
        textColor=colors.HexColor("#333333"),
        alignment=TA_LEFT,
    )

    small_style = ParagraphStyle(
        "Small",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
        textColor=colors.HexColor("#4B5563"),
    )

    cell_style = ParagraphStyle(
        "CellStyle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.6,
        leading=9,
        textColor=colors.HexColor("#1F2937"),
    )

    cell_bold_style = ParagraphStyle(
        "CellBoldStyle",
        parent=cell_style,
        fontName="Helvetica-Bold",
    )

    kpi_value_style = ParagraphStyle(
        "KPIValue",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=16,
        leading=18,
        textColor=colors.HexColor("#111111"),
        alignment=TA_CENTER,
    )

    kpi_label_style = ParagraphStyle(
        "KPILabel",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
        textColor=colors.HexColor("#6B7280"),
        alignment=TA_CENTER,
    )

    # =========================
    # HELPERS
    # =========================
    def valor_seguro(v):
        return pdf_safe(str(v)) if v not in (None, "", "None") else "-"

    def prioridade_badge(prioridade):
        p = (prioridade or "").strip().lower()
        if p in ["alta", "crítica", "critica"]:
            bg = "#FDECEC"
            fg = "#B42318"
        elif p in ["média", "media"]:
            bg = "#FFF4DB"
            fg = "#9A6700"
        elif p in ["baixa"]:
            bg = "#EAF4FF"
            fg = "#175CD3"
        else:
            bg = "#F3F4F6"
            fg = "#374151"

        return Paragraph(
            f"""<para align="center">
                <font color="{fg}">
                    <b>{valor_seguro(prioridade)}</b>
                </font>
            </para>""",
            ParagraphStyle(
                "badge_prio",
                parent=cell_style,
                backColor=colors.HexColor(bg),
                borderPadding=(3, 5, 3),
                alignment=TA_CENTER
            )
        )

    def status_badge(status):
        s = (status or "").strip().lower()
        if s in ["concluída", "concluida", "fechada", "finalizada"]:
            bg = "#E8F7EE"
            fg = "#146C43"
        elif s in ["em andamento", "pendente", "aberta"]:
            bg = "#FFF4DB"
            fg = "#9A6700"
        elif s in ["crítica", "critica", "atrasada"]:
            bg = "#FDE8EA"
            fg = "#B42318"
        else:
            bg = "#EEF2F6"
            fg = "#344054"

        return Paragraph(
            f"""<para align="center">
                <font color="{fg}">
                    <b>{valor_seguro(status)}</b>
                </font>
            </para>""",
            ParagraphStyle(
                "badge_status",
                parent=cell_style,
                backColor=colors.HexColor(bg),
                borderPadding=(3, 5, 3),
                alignment=TA_CENTER
            )
        )

    def p(txt, style=cell_style):
        return Paragraph(valor_seguro(txt), style)

    def draw_header_footer(canvas, doc):
        canvas.saveState()

        page_width, page_height = landscape(A4)

        # Barra superior vermelha
        canvas.setFillColor(colors.HexColor("#D40511"))
        canvas.rect(0, page_height - 12 * mm, page_width, 12 * mm, fill=1, stroke=0)

        # Faixa fina amarela abaixo
        canvas.setFillColor(colors.HexColor("#FFCC00"))
        canvas.rect(0, page_height - 14 * mm, page_width, 2 * mm, fill=1, stroke=0)

        # Cabeçalho
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica-Bold", 11)
        canvas.drawString(12 * mm, page_height - 8.2 * mm, "DHL SECURITY")

        canvas.setFont("Helvetica", 8)
        canvas.drawRightString(
            page_width - 12 * mm,
            page_height - 8.2 * mm,
            f"Relatório gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        )

        # Rodapé
        canvas.setStrokeColor(colors.HexColor("#D1D5DB"))
        canvas.setLineWidth(0.4)
        canvas.line(10 * mm, 10 * mm, page_width - 10 * mm, 10 * mm)

        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#6B7280"))
        canvas.drawString(10 * mm, 6 * mm, "Livro de Ocorrência • Passagem de Turno")
        canvas.drawRightString(page_width - 10 * mm, 6 * mm, f"Página {canvas.getPageNumber()}")

        canvas.restoreState()

    # =========================
    # CONTEÚDO
    # =========================
    elementos = []

    elementos.append(Paragraph("LIVRO DE OCORRÊNCIAS DE PASSAGEM DE TURNO", title_style))
    elementos.append(Paragraph(
        "Relatório corporativo consolidado das ocorrências registradas no sistema.",
        subtitle_style
    ))
    elementos.append(Spacer(1, 4))

    # Bloco de filtros
    filtros_data = [
        ["Data inicial", valor_seguro(filtros.get("data_inicial"))],
        ["Data final", valor_seguro(filtros.get("data_final"))],
        ["Turno", valor_seguro(filtros.get("turno"))],
        ["Status", valor_seguro(filtros.get("status"))],
        ["Site", valor_seguro(filtros.get("site"))],
    ]

    filtro_table = Table(
        [[Paragraph("<b>FILTROS APLICADOS</b>", section_label_style), ""]] +
        [[Paragraph(f"<b>{k}</b>", info_style), Paragraph(v, info_style)] for k, v in filtros_data],
        colWidths=[45 * mm, 95 * mm]
    )
    filtro_table.setStyle(TableStyle([
        ("SPAN", (0, 0), (1, 0)),
        ("BACKGROUND", (0, 0), (1, 0), colors.HexColor("#D40511")),
        ("TEXTCOLOR", (0, 0), (1, 0), colors.white),
        ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#FAFAFA")),
        ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#D1D5DB")),
        ("INNERGRID", (0, 1), (-1, -1), 0.35, colors.HexColor("#E5E7EB")),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))

    # KPI
    total_ocorrencias = len(rows)

    kpi_table = Table([
        [Paragraph(str(total_ocorrencias), kpi_value_style)],
        [Paragraph("TOTAL DE OCORRÊNCIAS", kpi_label_style)]
    ], colWidths=[42 * mm])

    kpi_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FFF8DB")),
        ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#FFCC00")),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))

    resumo_bloco = Table(
        [[filtro_table, kpi_table]],
        colWidths=[145 * mm, 50 * mm]
    )
    resumo_bloco.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))

    elementos.append(KeepTogether(resumo_bloco))
    elementos.append(Spacer(1, 10))

    # =========================
    # TABELA PRINCIPAL
    # =========================
    data = [[
        Paragraph("<b>ID</b>", cell_bold_style),
        Paragraph("<b>DATA / HORA</b>", cell_bold_style),
        Paragraph("<b>SITE</b>", cell_bold_style),
        Paragraph("<b>TURNO</b>", cell_bold_style),
        Paragraph("<b>SETOR</b>", cell_bold_style),
        Paragraph("<b>TIPO</b>", cell_bold_style),
        Paragraph("<b>PRIORIDADE</b>", cell_bold_style),
        Paragraph("<b>STATUS</b>", cell_bold_style),
        Paragraph("<b>RESP. SAÍDA</b>", cell_bold_style),
        Paragraph("<b>RESP. ENTRADA</b>", cell_bold_style),
    ]]

    for r in rows:
        data.append([
            p(r.id),
            p(r.data_hora_registro.strftime("%d/%m/%Y %H:%M") if r.data_hora_registro else "-"),
            p(r.site),
            p(r.turno),
            p(r.setor),
            p(r.tipo_ocorrencia),
            prioridade_badge(r.prioridade),
            status_badge(r.status),
            p(r.responsavel_saida),
            p(r.responsavel_entrada),
        ])

    tabela = Table(
        data,
        repeatRows=1,
        colWidths=[
            12 * mm,   # ID
            28 * mm,   # Data/Hora
            23 * mm,   # Site
            18 * mm,   # Turno
            28 * mm,   # Setor
            42 * mm,   # Tipo
            26 * mm,   # Prioridade
            30 * mm,   # Status
            38 * mm,   # Resp Saída
            38 * mm,   # Resp Entrada
        ]
    )

    tabela.setStyle(TableStyle([
        # Cabeçalho
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#FFCC00")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),

        # Corpo
        ("BACKGROUND", (0, 1), (-1, -1), colors.white),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#FCFCFD")]),
        ("TEXTCOLOR", (0, 1), (-1, -1), colors.HexColor("#111827")),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), 7.5),
        ("VALIGN", (0, 1), (-1, -1), "MIDDLE"),

        # Bordas
        ("LINEBELOW", (0, 0), (-1, 0), 0.9, colors.HexColor("#D1A800")),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#DADDE1")),
        ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#C9CDD3")),

        # Padding
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),

        # Alinhamentos específicos
        ("ALIGN", (0, 1), (0, -1), "CENTER"),   # ID
        ("ALIGN", (1, 1), (1, -1), "CENTER"),   # Data/Hora
        ("ALIGN", (3, 1), (3, -1), "CENTER"),   # Turno
        ("ALIGN", (6, 1), (7, -1), "CENTER"),   # Prioridade/Status
    ]))

    elementos.append(tabela)

    # Construção do PDF
    doc.build(
        elementos,
        onFirstPage=draw_header_footer,
        onLaterPages=draw_header_footer
    )

    output.seek(0)
    nome_arquivo = f"livro_ocorrencias_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"

    return send_file(
        output,
        as_attachment=True,
        download_name=nome_arquivo,
        mimetype="application/pdf"
    )

# =========================
# SETUP
# =========================
@app.route("/criar-banco")
def criar_banco():
    db.create_all()

    admin = User.query.filter_by(email="admin@dhl.com").first()
    if not admin:
        admin = User(
            nome="Administrador",
            email="admin@dhl.com",
            role="ADMIN",
            site="PG",
            is_active=True
        )
        admin.set_password("123456")
        db.session.add(admin)
        db.session.commit()

    return "Banco criado com sucesso. Login padrão: admin@dhl.com / 123456"


if __name__ == "__main__":
    with app.app_context():
        db.create_all()

        admin = User.query.filter_by(email="admin@dhl.com").first()
        if not admin:
            admin = User(
                nome="Administrador",
                email="admin@dhl.com",
                role="ADMIN",
                site="PG",
                is_active=True
            )
            admin.set_password("123456")
            db.session.add(admin)
            db.session.commit()

    app.run(debug=True)