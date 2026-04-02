import os
from io import BytesIO
from datetime import datetime, date
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, session, send_file
)
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func
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


app = Flask(__name__)
app.config["SECRET_KEY"] = "dev-secret-change-me"

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, "livro_ocorrencias.db")
app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{DB_PATH}"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)


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
            "acoes_tomadas": self.acoes_tomadas or "",
            "pendencias": self.pendencias or "",
            "status": self.status,
            "criado_por": self.criado_por or "",
        }


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
def parse_date_or_none(value: str):
    value = (value or "").strip()
    if not value:
        return None
    return datetime.strptime(value, "%Y-%m-%d").date()


def parse_datetime_local(value: str):
    value = (value or "").strip()
    if not value:
        return None
    return datetime.strptime(value, "%Y-%m-%dT%H:%M")


def get_filtros_ocorrencias():
    data_inicial = (request.args.get("data_inicial") or "").strip()
    data_final = (request.args.get("data_final") or "").strip()
    turno = (request.args.get("turno") or "").strip().upper()
    status = (request.args.get("status") or "").strip().upper()
    site = (request.args.get("site") or "").strip().upper()

    query = OcorrenciaTurno.query

    if data_inicial:
        di = datetime.strptime(data_inicial, "%Y-%m-%d").date()
        query = query.filter(OcorrenciaTurno.data_ocorrencia >= di)

    if data_final:
        df = datetime.strptime(data_final, "%Y-%m-%d").date()
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
# FECHAR OCORRÊNCIA
# =========================
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


# =========================
# CRIAÇÃO DE USUÁRIO PELA TELA DE LOGIN
# =========================
@app.route("/criar-usuario", methods=["GET", "POST"])
def criar_usuario():
    admin_publico_liberado = pode_criar_admin_publicamente()
    usuario_logado_admin = session.get("user_role") == "ADMIN"

    if request.method == "POST":
        nome = (request.form.get("nome") or "").strip()
        email = (request.form.get("email") or "").strip().lower()
        senha = request.form.get("senha") or ""
        role_solicitado = (request.form.get("role") or "USER").strip().upper()
        site = (request.form.get("site") or "").strip().upper()

        if not nome or not email or not senha:
            flash("Preencha nome, e-mail e senha.", "danger")
            return redirect(url_for("criar_usuario"))

        if role_solicitado not in ["ADMIN", "USER"]:
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
        role = (request.form.get("role") or "USER").strip().upper()
        site = (request.form.get("site") or "").strip().upper()

        if not nome or not email or not senha:
            flash("Preencha nome, e-mail e senha.", "danger")
            return redirect(url_for("novo_usuario"))

        if role not in ["ADMIN", "USER"]:
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
        role = (request.form.get("role") or "USER").strip().upper()
        site = (request.form.get("site") or "").strip().upper()
        is_active = (request.form.get("is_active") or "").strip() == "1"

        if not nome or not email:
            flash("Preencha nome e e-mail.", "danger")
            return redirect(url_for("editar_usuario", user_id=user_id))

        if role not in ["ADMIN", "USER"]:
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

        site = (request.form.get("site") or "").strip().upper()
        turno = (request.form.get("turno") or "").strip().upper()
        setor = (request.form.get("setor") or "").strip().upper()
        tipo_ocorrencia = (request.form.get("tipo_ocorrencia") or "").strip().upper()
        prioridade = (request.form.get("prioridade") or "").strip().upper()
        responsavel_saida = (request.form.get("responsavel_saida") or "").strip()
        responsavel_entrada = (request.form.get("responsavel_entrada") or "").strip()
        descricao = (request.form.get("descricao") or "").strip()
        acoes_tomadas = (request.form.get("acoes_tomadas") or "").strip()
        pendencias = (request.form.get("pendencias") or "").strip()
        status = (request.form.get("status") or "").strip().upper()

        if not all([
            data_ocorrencia, data_hora_registro, site, turno, setor,
            tipo_ocorrencia, prioridade, responsavel_saida,
            responsavel_entrada, descricao, status
        ]):
            flash("Preencha todos os campos obrigatórios.", "danger")
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

            ocorrencia.site = (request.form.get("site") or "").strip().upper()
            ocorrencia.turno = (request.form.get("turno") or "").strip().upper()
            ocorrencia.setor = (request.form.get("setor") or "").strip().upper()
            ocorrencia.tipo_ocorrencia = (request.form.get("tipo_ocorrencia") or "").strip().upper()
            ocorrencia.prioridade = (request.form.get("prioridade") or "").strip().upper()
            ocorrencia.responsavel_saida = (request.form.get("responsavel_saida") or "").strip()
            ocorrencia.responsavel_entrada = (request.form.get("responsavel_entrada") or "").strip()
            ocorrencia.descricao = (request.form.get("descricao") or "").strip()
            ocorrencia.acoes_tomadas = (request.form.get("acoes_tomadas") or "").strip() or None
            ocorrencia.pendencias = (request.form.get("pendencias") or "").strip() or None
            ocorrencia.status = (request.form.get("status") or "").strip().upper()
            ocorrencia.updated_at = datetime.now()

            if not all([
                ocorrencia.data_ocorrencia, ocorrencia.data_hora_registro, ocorrencia.site,
                ocorrencia.turno, ocorrencia.setor, ocorrencia.tipo_ocorrencia,
                ocorrencia.prioridade, ocorrencia.responsavel_saida,
                ocorrencia.responsavel_entrada, ocorrencia.descricao, ocorrencia.status
            ]):
                flash("Preencha todos os campos obrigatórios.", "danger")
                return redirect(url_for("editar_ocorrencia", ocorrencia_id=ocorrencia.id))

            db.session.commit()
            flash("Ocorrência atualizada com sucesso.", "success")
            return redirect(url_for("index"))

        except Exception as e:
            db.session.rollback()
            flash(f"Erro ao atualizar ocorrência: {e}", "danger")
            return redirect(url_for("editar_ocorrencia", ocorrencia_id=ocorrencia.id))

    return render_template("ocorrencia_form.html", ocorrencia=ocorrencia)


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
    em_aberto = db.session.query(func.count(OcorrenciaTurno.id)).filter(OcorrenciaTurno.status == "EM ABERTO").scalar() or 0
    acompanhamento = db.session.query(func.count(OcorrenciaTurno.id)).filter(OcorrenciaTurno.status == "EM ACOMPANHAMENTO").scalar() or 0
    finalizado = db.session.query(func.count(OcorrenciaTurno.id)).filter(OcorrenciaTurno.status == "FINALIZADO").scalar() or 0

    por_turno = (
        db.session.query(OcorrenciaTurno.turno, func.count(OcorrenciaTurno.id))
        .group_by(OcorrenciaTurno.turno)
        .order_by(OcorrenciaTurno.turno.asc())
        .all()
    )

    por_prioridade = (
        db.session.query(OcorrenciaTurno.prioridade, func.count(OcorrenciaTurno.id))
        .group_by(OcorrenciaTurno.prioridade)
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
        "Descrição", "Ações Tomadas", "Pendências", "Status",
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
            r.descricao,
            r.acoes_tomadas or "",
            r.pendencias or "",
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


@app.route("/ocorrencias/<int:ocorrencia_id>/pdf")
@login_required
def export_ocorrencia_individual_pdf(ocorrencia_id):
    ocorrencia = OcorrenciaTurno.query.get_or_404(ocorrencia_id)

    output = BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=A4,
        leftMargin=15 * mm,
        rightMargin=15 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "TituloDHL",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=18,
        textColor=colors.HexColor("#D40511"),
        spaceAfter=6,
        leading=22,
    )

    subtitle_style = ParagraphStyle(
        "SubTituloDHL",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        textColor=colors.HexColor("#555555"),
        spaceAfter=10,
    )

    label_style = ParagraphStyle(
        "Label",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10,
        textColor=colors.HexColor("#111111"),
        spaceAfter=2,
    )

    text_style = ParagraphStyle(
        "Texto",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10,
        leading=14,
        spaceAfter=8,
        textColor=colors.HexColor("#333333"),
    )

    box_style = TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FFFFFF")),
        ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#E0C24D")),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#F0E0A0")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ])

    elementos = []

    logo_path = os.path.join(app.static_folder, "logo.png")
    if os.path.exists(logo_path):
        logo = Image(logo_path, width=40 * mm, height=15 * mm)
        elementos.append(logo)
        elementos.append(Spacer(1, 5))

    elementos.append(Paragraph("DHL SECURITY - OCORRÊNCIA INDIVIDUAL", title_style))
    elementos.append(Paragraph("Relatório detalhado da ocorrência registrada", subtitle_style))

    dados_principais = [
        ["ID da ocorrência", f"#{ocorrencia.id}"],
        ["Data da ocorrência", ocorrencia.data_ocorrencia.strftime("%d/%m/%Y") if ocorrencia.data_ocorrencia else "-"],
        ["Data/Hora do registro", ocorrencia.data_hora_registro.strftime("%d/%m/%Y %H:%M") if ocorrencia.data_hora_registro else "-"],
        ["Site", ocorrencia.site or "-"],
        ["Turno", ocorrencia.turno or "-"],
        ["Setor", ocorrencia.setor or "-"],
        ["Tipo de ocorrência", ocorrencia.tipo_ocorrencia or "-"],
        ["Prioridade", ocorrencia.prioridade or "-"],
        ["Status", ocorrencia.status or "-"],
        ["Responsável saída", ocorrencia.responsavel_saida or "-"],
        ["Responsável entrada", ocorrencia.responsavel_entrada or "-"],
        ["Criado por", ocorrencia.criado_por or "-"],
        ["Criado em", ocorrencia.created_at.strftime("%d/%m/%Y %H:%M") if ocorrencia.created_at else "-"],
        ["Atualizado em", ocorrencia.updated_at.strftime("%d/%m/%Y %H:%M") if ocorrencia.updated_at else "-"],
    ]

    tabela_info = Table(dados_principais, colWidths=[50 * mm, 120 * mm])
    tabela_info.setStyle(box_style)
    elementos.append(tabela_info)
    elementos.append(Spacer(1, 10))

    elementos.append(Paragraph("Descrição:", label_style))
    elementos.append(Paragraph(ocorrencia.descricao or "-", text_style))

    elementos.append(Paragraph("Ações tomadas:", label_style))
    elementos.append(Paragraph(ocorrencia.acoes_tomadas or "-", text_style))

    elementos.append(Paragraph("Pendências:", label_style))
    elementos.append(Paragraph(ocorrencia.pendencias or "-", text_style))

    elementos.append(Spacer(1, 8))
    elementos.append(
        Paragraph(
            f"Documento gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            subtitle_style
        )
    )

    doc.build(elementos)

    output.seek(0)
    nome_arquivo = f"ocorrencia_{ocorrencia.id}.pdf"

    return send_file(
        output,
        as_attachment=True,
        download_name=nome_arquivo,
        mimetype="application/pdf"
    )


# =========================
# EXPORTAÇÃO PDF
# =========================
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
        topMargin=10 * mm,
        bottomMargin=10 * mm
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "DHLTitle",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=18,
        textColor=colors.HexColor("#D40511"),
        spaceAfter=8,
    )
    small_style = ParagraphStyle(
        "Small",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
    )

    elementos = []
    elementos.append(Paragraph("DHL SECURITY - LIVRO DE OCORRÊNCIA DE PASSAGEM DE TURNO", title_style))

    filtro_txt = (
        f"Filtros aplicados | Data inicial: {filtros['data_inicial'] or '-'} | "
        f"Data final: {filtros['data_final'] or '-'} | Turno: {filtros['turno'] or '-'} | "
        f"Status: {filtros['status'] or '-'} | Site: {filtros['site'] or '-'}"
    )
    elementos.append(Paragraph(filtro_txt, small_style))
    elementos.append(Spacer(1, 6))

    data = [[
        "ID", "Data/Hora", "Site", "Turno", "Setor", "Tipo",
        "Prioridade", "Status", "Resp. Saída", "Resp. Entrada"
    ]]

    for r in rows:
        data.append([
            str(r.id),
            r.data_hora_registro.strftime("%d/%m/%Y %H:%M") if r.data_hora_registro else "",
            r.site,
            r.turno,
            r.setor,
            r.tipo_ocorrencia,
            r.prioridade,
            r.status,
            r.responsavel_saida,
            r.responsavel_entrada,
        ])

    tabela = Table(data, repeatRows=1)
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#FFCC00")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CCCCCC")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#FFF9E6")]),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))

    elementos.append(tabela)
    doc.build(elementos)

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