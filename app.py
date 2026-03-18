import streamlit as st
import openpyxl
from datetime import datetime, timedelta, date
import math
import json
import hashlib
import smtplib
import os
import random
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from zoneinfo import ZoneInfo

# ====================== CONFIG EMAIL ======================
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_EMAIL = "rlnt.gestion@gmail.com"
SMTP_PASSWORD = st.secrets["smtp"]["password"]
ADMIN_EMAIL = "rlnt.gestion@gmail.com"
USERS_FILE = "users.json"
APP_URL = "https://gestion-rl-nt.streamlit.app/"   # ←←← CHANGE ÇA AVEC TON VRAI LIEN !

def load_users():
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {
        ADMIN_EMAIL: {
            "password": hashlib.sha256("admin123".encode()).hexdigest(),
            "role": "Admin",
            "name": "Administrateur"
        }
    }

def save_users(users):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

def hash_password(pw):
    return hashlib.sha256(pw.encode()).hexdigest()

def generate_temp_password():
    return f"temp{random.randint(1000,9999)}"

def send_email(to_email, subject, body, attachment_path=None):
    try:
        msg = MIMEMultipart()
        msg["From"] = SMTP_EMAIL
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))
        if attachment_path and os.path.exists(attachment_path):
            with open(attachment_path, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(attachment_path)}"')
                msg.attach(part)
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SMTP_EMAIL, SMTP_PASSWORD)
        server.sendmail(SMTP_EMAIL, to_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"Erreur email : {str(e)}")
        return False

def send_to_all_users(subject, body, attachment_path=None):
    users = load_users()
    for email in users:
        send_email(email, subject, body, attachment_path)

# ====================== LOGIN ======================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.email = None

if not st.session_state.logged_in:
    st.title("🔐 Connexion - Gestion Contrats RL/NT")
    email = st.text_input("Email", value=ADMIN_EMAIL)
    password = st.text_input("Mot de passe", type="password", value="admin123")
    if st.button("Se connecter"):
        users = load_users()
        if email in users and users[email]["password"] == hash_password(password):
            st.session_state.logged_in = True
            st.session_state.role = users[email]["role"]
            st.session_state.email = email
            st.rerun()
        else:
            st.error("❌ Identifiants incorrects")
    st.stop()

# ====================== PANNEAU ADMINISTRATEUR ======================
if st.session_state.role == "Admin":
    st.subheader("👑 Panneau Administrateur - Gestion des utilisateurs")
    with st.expander("➕ Ajouter un nouvel utilisateur", expanded=True):
        with st.form("add_user_form"):
            new_email = st.text_input("Email du nouvel utilisateur")
            new_role = st.selectbox("Rôle", ["RL", "NT", "Admin"])
            new_name = st.text_input("Nom complet")
            if st.form_submit_button("Ajouter l'utilisateur"):
                if new_email and new_name:
                    users = load_users()
                    if new_email in users:
                        st.error("❌ Cet email existe déjà")
                    else:
                        temp_pw = generate_temp_password()
                        users[new_email] = {
                            "password": hash_password(temp_pw),
                            "role": new_role,
                            "name": new_name
                        }
                        save_users(users)
                        email_body_new = f"""Bonjour {new_name},

Voici tes identifiants temporaires :
Email : {new_email}
Mot de passe : {temp_pw}

Accède directement à l'application ici :
{APP_URL}

Tu devras changer ce mot de passe dès ta première connexion.

Cordialement,
L'équipe RL/NT"""
                        send_email(new_email, "Bienvenue - Accès RL/NT", email_body_new)
                        send_email(ADMIN_EMAIL, "Nouvel utilisateur ajouté + users.json", f"{new_name} ({new_email} - {new_role}) ajouté.", USERS_FILE)
                        st.success(f"✅ {new_name} ajouté avec succès !")
                        st.info(f"**Mot de passe temporaire pour {new_email} : {temp_pw}**")
                        st.warning("⚠️ COPIE-LE MAINTENANT !")
                        st.rerun()
                else:
                    st.error("Veuillez remplir tous les champs")

    st.subheader("Utilisateurs existants")
    users = load_users()
    for email, data in list(users.items()):
        col1, col2, col3 = st.columns([3, 2, 1])
        with col1:
            st.write(f"**{data['name']}** - {email} ({data['role']})")
        with col2:
            st.write("✅ Actif")
        with col3:
            if email != ADMIN_EMAIL and st.button("🗑 Supprimer", key=f"del_{email}"):
                del users[email]
                save_users(users)
                st.success(f"{email} supprimé")
                st.rerun()

# ====================== SIDEBAR ======================
with st.sidebar:
    st.header("🔑 Mon compte")
    st.write(f"Connecté : **{st.session_state.email}**")
    st.write(f"Rôle : **{st.session_state.role}**")
    with st.expander("Changer mon mot de passe"):
        with st.form("change_password_form"):
            old_pw = st.text_input("Ancien mot de passe", type="password")
            new_pw = st.text_input("Nouveau mot de passe", type="password")
            confirm_pw = st.text_input("Confirmer", type="password")
            if st.form_submit_button("Changer"):
                users = load_users()
                if hash_password(old_pw) == users[st.session_state.email]["password"]:
                    if new_pw == confirm_pw and len(new_pw) >= 6:
                        users[st.session_state.email]["password"] = hash_password(new_pw)
                        save_users(users)
                        send_email(ADMIN_EMAIL, "🔄 Mise à jour users.json", f"Mot de passe modifié par {st.session_state.email}.", USERS_FILE)
                        st.success("✅ Mot de passe changé !")
                        st.rerun()
                    else:
                        st.error("❌ Mots de passe ne correspondent pas ou trop courts")
                else:
                    st.error("❌ Ancien mot de passe incorrect.")

# ====================== FONCTIONS ======================
FRENCH_MONTHS = {1: "Janvier", 2: "Février", 3: "Mars", 4: "Avril", 5: "Mai", 6: "Juin", 7: "Juillet", 8: "Août", 9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Décembre"}

def safe_float(val):
    if val is None: return 0.0
    try: return float(val)
    except: return 0.0

def get_previous_monday():
    today = datetime.today().date()
    return today - timedelta(days=today.weekday())

def normalize_date(d):
    if isinstance(d, datetime): return d.date()
    elif isinstance(d, date): return d
    elif isinstance(d, (int, float)) and d > 40000:
        try: return (datetime(1899, 12, 30) + timedelta(days=int(d))).date()
        except: return None
    elif isinstance(d, str):
        try: return datetime.strptime(d[:10], "%Y-%m-%d").date()
        except: return None
    return None

def get_monday_of_week(selected_date):
    if isinstance(selected_date, datetime): d = selected_date.date()
    else: d = selected_date
    return d - timedelta(days=d.weekday())

def find_project_row(ws, project_name, start_row=3):
    project_name = str(project_name).strip().lower()
    for r in range(start_row, ws.max_row + 200):
        if str(ws.cell(r, 1).value or "").strip().lower() == project_name:
            return r
    return None

def get_project_status(ws_desc, proj_name):
    proj_name = str(proj_name).strip()
    for r in range(3, ws_desc.max_row + 1):
        if str(ws_desc.cell(r, 1).value or "").strip() == proj_name:
            return str(ws_desc.cell(r, 2).value or "En soumission").strip()
    return "En soumission"

def get_display_name(proj_name, status):
    return f"{proj_name} - {status}"

def find_or_create_week_column(ws_cal, monday_date):
    target_date = normalize_date(monday_date)
    for c in range(2, ws_cal.max_column + 1):
        if normalize_date(ws_cal.cell(4, c).value) == target_date:
            return c
    col = ws_cal.max_column + 1
    cell = ws_cal.cell(4, col, monday_date)
    cell.number_format = "dd/mm/yyyy"
    cell.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90)
    ws_cal.column_dimensions[get_column_letter(col)].width = 5.0
    return col

def find_last_filled_column(ws_cal, proj_row, before_col):
    for c in range(before_col - 1, 1, -1):
        if safe_float(ws_cal.cell(proj_row + 9, c).value) > 0:
            return c
    return None

def backfill_intermediate_weeks(ws_cal, proj_row, last_col, new_col):
    if not last_col or new_col <= last_col + 1: return
    for c in range(last_col + 1, new_col):
        for offset in range(1, 10):
            ws_cal.cell(proj_row + offset, c, ws_cal.cell(proj_row + offset, last_col).value)

def find_last_used_column(ws):
    for c in range(ws.max_column, 1, -1):
        if ws.cell(4, c).value is not None:
            return c
    return 20

def save_gantt_data(ws_gantt):
    data = {}
    r = 5
    while r <= ws_gantt.max_row:
        val = str(ws_gantt.cell(r, 1).value or "").strip()
        if val and not val.startswith(("Besoin", "CAMPS", "SOUS-TOTAL", "TOTAL")):
            plain_name = val.split(" - ")[0] if " - " in val else val
            block_data = []
            for off in range(1, 5):
                row_data = [safe_float(ws_gantt.cell(r + off, c).value) for c in range(2, ws_gantt.max_column + 1)]
                block_data.append(row_data)
            data[plain_name] = block_data
            r += 6
            continue
        r += 1
    return data

def restore_gantt_data(ws_gantt, saved_data):
    r = 5
    while r <= ws_gantt.max_row:
        val = str(ws_gantt.cell(r, 1).value or "").strip()
        if val and not val.startswith(("Besoin", "CAMPS", "SOUS-TOTAL", "TOTAL")):
            plain_name = val.split(" - ")[0] if " - " in val else val
            if plain_name in saved_data:
                block_data = saved_data[plain_name]
                for off in range(1, 5):
                    for c_idx, v in enumerate(block_data[off-1], start=2):
                        ws_gantt.cell(r + off, c_idx, v)
            r += 6
            continue
        r += 1

def save_calendrier_data(ws_cal):
    data = {}
    r = 5
    while r <= ws_cal.max_row:
        val = str(ws_cal.cell(r, 1).value or "").strip()
        if val and not val.startswith(("Dortoir", "Bureau", "Vaste", "Total", "CAMPS", "TOTAL")):
            plain_name = val.split(" - ")[0] if " - " in val else val
            block_data = []
            for off in range(1, 10):
                row_data = [safe_float(ws_cal.cell(r + off, c).value) for c in range(2, ws_cal.max_column + 1)]
                block_data.append(row_data)
            data[plain_name] = block_data
            r += 11
            continue
        r += 1
    return data

def restore_calendrier_data(ws_cal, saved_data):
    r = 5
    while r <= ws_cal.max_row:
        val = str(ws_cal.cell(r, 1).value or "").strip()
        if val and not val.startswith(("Dortoir", "Bureau", "Vaste", "Total", "CAMPS", "TOTAL")):
            plain_name = val.split(" - ")[0] if " - " in val else val
            if plain_name in saved_data:
                block_data = saved_data[plain_name]
                for off in range(1, 10):
                    for c_idx, v in enumerate(block_data[off-1], start=2):
                        ws_cal.cell(r + off, c_idx, v)
            r += 11
            continue
        r += 1

def check_gantt_gaps(ws_gantt):
    warnings = []
    last_col = find_last_used_column(ws_gantt)
    r = 5
    while r <= ws_gantt.max_row:
        val = str(ws_gantt.cell(r, 1).value or "").strip()
        if val and not val.startswith(("Besoin", "CAMPS", "SOUS-TOTAL", "TOTAL")):
            proj_name = val.split(" - ")[0] if " - " in val else val
            last_filled = None
            for c in range(2, last_col + 1):
                if safe_float(ws_gantt.cell(r + 1, c).value) > 0:
                    if last_filled and c > last_filled + 1:
                        gap = c - last_filled - 1
                        start = get_week_date(ws_gantt, last_filled + 1)
                        end = get_week_date(ws_gantt, c - 1)
                        warnings.append(f"**{proj_name}** : {gap} semaine(s) vide(s) du **{start}** au **{end}**")
                    last_filled = c
            r += 6
            continue
        r += 1
    return warnings

def get_week_date(ws, col):
    val = ws.cell(4, col).value
    dt = normalize_date(val)
    return dt.strftime("%d/%m/%Y") if dt else f"col {col}"

def remove_existing_total_and_blocks(ws):
    r = ws.max_row
    while r >= 1:
        val = str(ws.cell(r, 1).value or "").strip().upper()
        if val in ("TOTAL", "SOUS-TOTAL - CONTRAT OBTENU", "SOUS-TOTAL - SOUMISSION"):
            ws.delete_rows(r, 20)
            r = ws.max_row
            continue
        r -= 1
    r = 5
    while r <= ws.max_row:
        val = str(ws.cell(r, 1).value or "").strip()
        if val and not val.startswith(("Besoin", "CAMPS", "SOUS-TOTAL", "TOTAL")):
            ws.delete_rows(r, 6)
            continue
        r += 1

def rebuild_gantt_sheet(ws_gantt, ws_desc, projects):
    saved_data = save_gantt_data(ws_gantt)
    remove_existing_total_and_blocks(ws_gantt)
    status_dict = {proj: get_project_status(ws_desc, proj) for proj in projects}
    obtained = []
    soumission = []
    for proj in projects:
        status = status_dict.get(proj, "En soumission")
        if status == "Abandonné": continue
        display_name = get_display_name(proj, status)
        if status == "Contrat obtenu":
            obtained.append((proj, display_name))
        else:
            soumission.append((proj, display_name))
    current_row = 5
    obtained_start_rows = []
    soumission_start_rows = []
    all_start_rows = []
    for proj_plain, display_name in obtained + soumission:
        ws_gantt.cell(current_row, 1, display_name)
        ws_gantt.cell(current_row + 1, 1, "Besoin Lit")
        ws_gantt.cell(current_row + 2, 1, "Besoin dortoi")
        ws_gantt.cell(current_row + 3, 1, "Besoin moudule bureau")
        ws_gantt.cell(current_row + 4, 1, "Besoin module vaste")
        if status_dict.get(proj_plain) == "Contrat obtenu":
            obtained_start_rows.append(current_row)
        else:
            soumission_start_rows.append(current_row)
        all_start_rows.append(current_row)
        current_row += 6
    last_col = find_last_used_column(ws_gantt)
    restore_gantt_data(ws_gantt, saved_data)
    # SOUS-TOTAL - Contrat obtenu
    sub_obt_row = current_row
    ws_gantt.cell(sub_obt_row, 1, "SOUS-TOTAL - Contrat obtenu")
    ws_gantt.cell(sub_obt_row + 1, 1, "Total Besoin Lit")
    ws_gantt.cell(sub_obt_row + 2, 1, "Total Besoin dortoi")
    ws_gantt.cell(sub_obt_row + 3, 1, "Total Besoin moudule bureau")
    ws_gantt.cell(sub_obt_row + 4, 1, "Total Besoin module vaste")
    for c in range(2, last_col + 1):
        lit = sum(safe_float(ws_gantt.cell(start_r + 1, c).value) for start_r in obtained_start_rows)
        dort = sum(safe_float(ws_gantt.cell(start_r + 2, c).value) for start_r in obtained_start_rows)
        bur = sum(safe_float(ws_gantt.cell(start_r + 3, c).value) for start_r in obtained_start_rows)
        vas = sum(safe_float(ws_gantt.cell(start_r + 4, c).value) for start_r in obtained_start_rows)
        ws_gantt.cell(sub_obt_row + 1, c, lit)
        ws_gantt.cell(sub_obt_row + 2, c, dort)
        ws_gantt.cell(sub_obt_row + 3, c, bur)
        ws_gantt.cell(sub_obt_row + 4, c, vas)
    # SOUS-TOTAL - Soumission
    sub_sou_row = sub_obt_row + 6
    ws_gantt.cell(sub_sou_row, 1, "SOUS-TOTAL - Soumission")
    ws_gantt.cell(sub_sou_row + 1, 1, "Total Besoin Lit")
    ws_gantt.cell(sub_sou_row + 2, 1, "Total Besoin dortoi")
    ws_gantt.cell(sub_sou_row + 3, 1, "Total Besoin moudule bureau")
    ws_gantt.cell(sub_sou_row + 4, 1, "Total Besoin module vaste")
    for c in range(2, last_col + 1):
        lit = sum(safe_float(ws_gantt.cell(start_r + 1, c).value) for start_r in soumission_start_rows)
        dort = sum(safe_float(ws_gantt.cell(start_r + 2, c).value) for start_r in soumission_start_rows)
        bur = sum(safe_float(ws_gantt.cell(start_r + 3, c).value) for start_r in soumission_start_rows)
        vas = sum(safe_float(ws_gantt.cell(start_r + 4, c).value) for start_r in soumission_start_rows)
        ws_gantt.cell(sub_sou_row + 1, c, lit)
        ws_gantt.cell(sub_sou_row + 2, c, dort)
        ws_gantt.cell(sub_sou_row + 3, c, bur)
        ws_gantt.cell(sub_sou_row + 4, c, vas)
    # TOTAL GÉNÉRAL
    total_row = sub_sou_row + 6
    ws_gantt.cell(total_row, 1, "TOTAL")
    ws_gantt.cell(total_row + 1, 1, "Total Besoin Lit")
    ws_gantt.cell(total_row + 2, 1, "Total Besoin dortoi")
    ws_gantt.cell(total_row + 3, 1, "Total Besoin moudule bureau")
    ws_gantt.cell(total_row + 4, 1, "Total Besoin module vaste")
    for c in range(2, last_col + 1):
        lit = sum(safe_float(ws_gantt.cell(start_r + 1, c).value) for start_r in all_start_rows)
        dort = sum(safe_float(ws_gantt.cell(start_r + 2, c).value) for start_r in all_start_rows)
        bur = sum(safe_float(ws_gantt.cell(start_r + 3, c).value) for start_r in all_start_rows)
        vas = sum(safe_float(ws_gantt.cell(start_r + 4, c).value) for start_r in all_start_rows)
        ws_gantt.cell(total_row + 1, c, lit)
        ws_gantt.cell(total_row + 2, c, dort)
        ws_gantt.cell(total_row + 3, c, bur)
        ws_gantt.cell(total_row + 4, c, vas)

def remove_existing_calendrier_blocks_and_total(ws_cal):
    r = ws_cal.max_row
    while r >= 1:
        val = str(ws_cal.cell(r, 1).value or "").strip().upper()
        if val == "TOTAL":
            ws_cal.delete_rows(r, 20)
            r = ws_cal.max_row
            continue
        r -= 1
    r = 5
    while r <= ws_cal.max_row:
        val = str(ws_cal.cell(r, 1).value or "").strip()
        if val and not val.startswith(("Dortoir", "Bureau", "Vaste", "Total", "CAMPS")):
            ws_cal.delete_rows(r, 11)
            continue
        r += 1

def rebuild_calendrier_sheet(ws_cal, ws_desc, projects):
    saved_data = save_calendrier_data(ws_cal)
    remove_existing_calendrier_blocks_and_total(ws_cal)
    current_row = 5
    project_start_rows = []
    for proj in projects:
        if get_project_status(ws_desc, proj) != "Contrat obtenu":
            continue
        ws_cal.cell(current_row, 1, proj)
        ws_cal.cell(current_row + 1, 1, "Dortoir RL")
        ws_cal.cell(current_row + 2, 1, "Bureau RL")
        ws_cal.cell(current_row + 3, 1, "Vaste RL")
        ws_cal.cell(current_row + 4, 1, "Total RL")
        ws_cal.cell(current_row + 5, 1, "Dortoir NT")
        ws_cal.cell(current_row + 6, 1, "Bureau NT")
        ws_cal.cell(current_row + 7, 1, "Vaste NT")
        ws_cal.cell(current_row + 8, 1, "Total NT")
        ws_cal.cell(current_row + 9, 1, f"Total Module RL projet {proj}")
        project_start_rows.append(current_row)
        current_row += 11
    total_row = current_row + 1
    ws_cal.cell(total_row, 1, "TOTAL")
    labels = ["Dortoir RL", "Bureau RL", "Vaste RL", "Total RL", "Dortoir NT", "Bureau NT", "Vaste NT", "Total NT", "Total Module RL projet TOTAL"]
    for i, lbl in enumerate(labels):
        ws_cal.cell(total_row + i + 1, 1, lbl)
    last_col = find_last_used_column(ws_cal)
    restore_calendrier_data(ws_cal, saved_data)
    for c in range(2, last_col + 1):
        dort_rl = bur_rl = vas_rl = dort_nt = bur_nt = vas_nt = 0.0
        for start_r in project_start_rows:
            dort_rl += safe_float(ws_cal.cell(start_r + 1, c).value)
            bur_rl += safe_float(ws_cal.cell(start_r + 2, c).value)
            vas_rl += safe_float(ws_cal.cell(start_r + 3, c).value)
            dort_nt += safe_float(ws_cal.cell(start_r + 5, c).value)
            bur_nt += safe_float(ws_cal.cell(start_r + 6, c).value)
            vas_nt += safe_float(ws_cal.cell(start_r + 7, c).value)
        total_rl = dort_rl + bur_rl + vas_rl
        total_nt = dort_nt + bur_nt + vas_nt
        total_module = total_rl + total_nt
        ws_cal.cell(total_row + 1, c, dort_rl)
        ws_cal.cell(total_row + 2, c, bur_rl)
        ws_cal.cell(total_row + 3, c, vas_rl)
        ws_cal.cell(total_row + 4, c, total_rl)
        ws_cal.cell(total_row + 5, c, dort_nt)
        ws_cal.cell(total_row + 6, c, bur_nt)
        ws_cal.cell(total_row + 7, c, vas_nt)
        ws_cal.cell(total_row + 8, c, total_nt)
        ws_cal.cell(total_row + 9, c, total_module)

# ====================== STYLING & RATTRAPAGE ======================
def apply_thin_grid(ws, min_row, max_row, min_col, max_col):
    thin_side = Side(style="thin", color="000000")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            ws.cell(r, c).border = thin_border

def create_combined_border(top_thick=False, bottom_thick=False):
    thin_side = Side(style="thin", color="000000")
    thick_side = Side(style="thick", color="000000")
    return Border(left=thin_side, right=thin_side, top=thick_side if top_thick else thin_side, bottom=thick_side if bottom_thick else thin_side)

def apply_month_headers(ws):
    if ws.title not in ["Gantt Besoins", "Calendrier réel"]:
        return
    last_col = find_last_used_column(ws)
    for merged in list(ws.merged_cells.ranges):
        if merged.min_row == 3 and merged.max_row == 3:
            ws.unmerge_cells(str(merged))
    for c in range(1, last_col + 5):
        ws.cell(3, c).value = None
    start_col = None
    current_key = None
    current_name = None
    for c in range(2, last_col + 2):
        dt = normalize_date(ws.cell(4, c).value) if c <= last_col else None
        new_key = (dt.year, dt.month) if dt else None
        new_name = f"{FRENCH_MONTHS[dt.month]} {dt.year}" if dt else None
        if new_key != current_key:
            if start_col is not None and c > start_col:
                end_col = c - 1
                ws.merge_cells(start_row=3, start_column=start_col, end_row=3, end_column=end_col)
                cell = ws.cell(3, start_col, current_name)
                cell.font = Font(bold=True, size=12)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            current_key = new_key
            start_col = c
            current_name = new_name
    if start_col is not None and current_key:
        ws.merge_cells(start_row=3, start_column=start_col, end_row=3, end_column=last_col)
        cell = ws.cell(3, start_col, current_name)
        cell.font = Font(bold=True, size=12)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 30

def apply_all_styling(wb):
    center_align = Alignment(horizontal="center", vertical="center")
    vertical_date = Alignment(horizontal="center", vertical="center", text_rotation=90)
    bold_font = Font(bold=True)
    thin_side = Side(style="thin", color="000000")
    thick_side = Side(style="thick", color="000000")
    timestamp = datetime.now(ZoneInfo("America/Montreal")).strftime("%d/%m/%Y %H:%M")
    version_text = f"Version du {timestamp}"
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if sheet_name == "Description projet et engag. RL":
            last_col = 25
        elif sheet_name in ["Gantt Besoins", "Calendrier réel"]:
            last_col = find_last_used_column(ws)
        else:
            last_col = 25
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=last_col):
            for cell in row:
                if cell.value is not None:
                    cell.alignment = center_align
        if sheet_name in ["Gantt Besoins", "Calendrier réel"]:
            for c in range(2, last_col + 1):
                if ws.cell(4, c).value:
                    ws.cell(4, c).alignment = vertical_date
            apply_month_headers(ws)
        if sheet_name == "Description projet et engag. RL":
            apply_thin_grid(ws, 5, ws.max_row, 1, 25)
            section_starts = [5, 9, 13, 17, 21]
            for r in range(5, ws.max_row + 1):
                for col in section_starts:
                    cell = ws.cell(r, col)
                    border = cell.border
                    cell.border = Border(left=thick_side, right=border.right, top=border.top, bottom=border.bottom)
                    if col > 1:
                        prev_cell = ws.cell(r, col - 1)
                        border_prev = prev_cell.border
                        prev_cell.border = Border(right=thick_side, left=border_prev.left, top=border_prev.top, bottom=border_prev.bottom)
            for r in range(5, ws.max_row + 1):
                cell = ws.cell(r, 25)
                border = cell.border
                cell.border = Border(right=thick_side, left=border.left, top=border.top, bottom=border.bottom)
            for c in range(1, 26):
                ws.cell(1, c).font = bold_font
                ws.cell(2, c).font = bold_font
            ws.row_dimensions[5].height = 110
            wrap_align = Alignment(wrap_text=True, horizontal="center", vertical="center")
            for c in range(1, 26):
                cell = ws.cell(5, c)
                cell.font = bold_font
                cell.alignment = wrap_align
            for r in range(6, ws.max_row + 1):
                for c in range(1, 26):
                    ws.cell(r, c).alignment = Alignment(horizontal="center", vertical="center")
        elif sheet_name == "Gantt Besoins":
            r = 5
            while r <= ws.max_row:
                val = str(ws.cell(r, 1).value or "").strip()
                if val and not val.startswith(("Besoin", "CAMPS")) or "SOUS-TOTAL" in val.upper() or val.upper() == "TOTAL":
                    ws.cell(r, 1).font = bold_font
                    apply_thin_grid(ws, r + 1, r + 4, 1, last_col)
                    for c in range(1, last_col + 1):
                        ws.cell(r, c).border = Border(bottom=thick_side)
                    for c in range(1, last_col + 1):
                        ws.cell(r + 4, c).border = create_combined_border(bottom_thick=True)
                    r += 6
                    continue
                r += 1
        elif sheet_name == "Calendrier réel":
            r = 5
            while r <= ws.max_row:
                val = str(ws.cell(r, 1).value or "").strip()
                if val and not val.startswith(("Dortoir", "Bureau", "Vaste", "Total", "CAMPS")) or val.upper() == "TOTAL":
                    ws.cell(r, 1).font = bold_font
                    apply_thin_grid(ws, r + 1, r + 9, 1, last_col)
                    for c in range(1, last_col + 1):
                        ws.cell(r, c).border = Border(bottom=thick_side)
                    for c in range(1, last_col + 1):
                        ws.cell(r + 4, c).border = Border(top=thick_side, bottom=thick_side, left=thin_side, right=thin_side)
                    for c in range(1, last_col + 1):
                        ws.cell(r + 8, c).border = Border(top=thick_side, bottom=thick_side, left=thin_side, right=thin_side)
                    for c in range(1, last_col + 1):
                        ws.cell(r + 9, c).border = create_combined_border(bottom_thick=True)
                    for offset in [4, 8, 9]:
                        ws.cell(r + offset, 1).font = bold_font
                        for c in range(2, last_col + 1):
                            ws.cell(r + offset, c).font = bold_font
                    r += 11
                    continue
                r += 1
        elif sheet_name == "Rattrapage":
            for c in range(1, 12):
                ws.cell(1, c).font = bold_font
                ws.cell(1, c).border = Border(bottom=thick_side)
            for r in range(2, ws.max_row + 1):
                for c in range(1, 12):
                    ws.cell(r, c).border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        if sheet_name != "Rattrapage":
            merge_col = 10 if sheet_name == "Calendrier réel" else 4
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=merge_col)
            ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=merge_col)
            ws.cell(1, 1).font = Font(bold=True, size=14)
            ws.cell(1, 1).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(2, 1, version_text).font = Font(bold=True, size=11)
            ws.cell(2, 1).alignment = Alignment(horizontal="center", vertical="center")
    if "Gantt Besoins" in wb.sheetnames:
        wb["Gantt Besoins"].freeze_panes = "B5"
    if "Calendrier réel" in wb.sheetnames:
        wb["Calendrier réel"].freeze_panes = "B5"
    if "Description projet et engag. RL" in wb.sheetnames:
        wb["Description projet et engag. RL"].freeze_panes = "F6"

def update_rattrapage_sheet(wb):
    ws_cal = wb['Calendrier réel']
    if 'Rattrapage' in wb.sheetnames:
        wb.remove(wb['Rattrapage'])
    ws_rat = wb.create_sheet("Rattrapage")
    headers = ["Projet", "Pic_Réel", "Max_RL", "Max_NT", "% RL Réel", "% NT Réel", "Base_NT_30%", "Rattrapage_Créé", "Rattrapage_Utilisé", "Cumul_Déficit_Global", "Note"]
    for c, header in enumerate(headers, 1):
        cell = ws_rat.cell(1, c, header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    project_data = []
    total_cree = total_utilise = 0.0
    for proj in st.session_state.projects:
        proj_row = find_project_row(ws_cal, proj, start_row=5)
        if not proj_row: continue
        pic_reel = max_rl = max_nt = 0.0
        for c in range(2, ws_cal.max_column + 1):
            if ws_cal.cell(4, c).value is not None:
                total_module = safe_float(ws_cal.cell(proj_row + 9, c).value)
                total_rl = safe_float(ws_cal.cell(proj_row + 4, c).value)
                total_nt = safe_float(ws_cal.cell(proj_row + 8, c).value)
                pic_reel = max(pic_reel, total_module)
                max_rl = max(max_rl, total_rl)
                max_nt = max(max_nt, total_nt)
        if pic_reel == 0: continue
        base_nt_30 = math.ceil(pic_reel * 0.3)
        rattrapage_cree = max(0.0, base_nt_30 - max_nt)
        rattrapage_utilise = max(0.0, max_nt - base_nt_30)
        total_cree += rattrapage_cree
        total_utilise += rattrapage_utilise
        pct_rl = round((max_rl / pic_reel * 100), 1) if pic_reel > 0 else 0
        pct_nt = round((max_nt / pic_reel * 100), 1) if pic_reel > 0 else 0
        project_data.append({
            'proj': proj, 'pic_reel': int(pic_reel), 'max_rl': int(max_rl), 'max_nt': int(max_nt),
            'pct_rl': f"{pct_rl}%", 'pct_nt': f"{pct_nt}%", 'base_nt_30': int(base_nt_30),
            'cree': int(rattrapage_cree), 'utilise': int(rattrapage_utilise)
        })
    cumul_global = total_cree - total_utilise
    write_row = 2
    for data in project_data:
        ws_rat.cell(write_row, 1, data['proj'])
        ws_rat.cell(write_row, 2, data['pic_reel'])
        ws_rat.cell(write_row, 3, data['max_rl'])
        ws_rat.cell(write_row, 4, data['max_nt'])
        ws_rat.cell(write_row, 5, data['pct_rl'])
        ws_rat.cell(write_row, 6, data['pct_nt'])
        ws_rat.cell(write_row, 7, data['base_nt_30'])
        ws_rat.cell(write_row, 8, data['cree'])
        ws_rat.cell(write_row, 9, data['utilise'])
        ws_rat.cell(write_row, 10, int(cumul_global))
        ws_rat.cell(write_row, 11, "Calcul automatique")
        write_row += 1
    if project_data:
        total_row = write_row
        ws_rat.cell(total_row, 1, "TOTAL GÉNÉRAL")
        ws_rat.cell(total_row, 8, int(total_cree))
        ws_rat.cell(total_row, 9, int(total_utilise))
        ws_rat.cell(total_row, 10, int(cumul_global))
        ws_rat.cell(total_row, 11, "Cumul Global")
        for c in [1, 8, 9, 10, 11]:
            ws_rat.cell(total_row, c).font = Font(bold=True)
    for c in range(1, 12):
        ws_rat.column_dimensions[get_column_letter(c)].width = 20
    ws_rat.freeze_panes = "B2"

# ====================== INTERFACE ======================
st.title("Gestion Contrats RL - Calendrier & Calculateur")
st.markdown(f"**Connecté en tant que : {st.session_state.email} ({st.session_state.role})**")

uploaded_file = st.file_uploader("Upload Excel (Modèle Base.xlsx)", type="xlsx")

if 'wb' not in st.session_state:
    st.session_state.wb = None
if 'projects' not in st.session_state:
    st.session_state.projects = []
if 'gantt_gap_confirmed' not in st.session_state:
    st.session_state.gantt_gap_confirmed = False
if 'initial_rebuild_done' not in st.session_state:
    st.session_state.initial_rebuild_done = False

if uploaded_file:
    if st.session_state.wb is None or st.button("Recharger le fichier"):
        st.session_state.wb = openpyxl.load_workbook(uploaded_file, data_only=False)
        st.session_state.initial_rebuild_done = False

    wb = st.session_state.wb
    ws_desc = wb['Description projet et engag. RL']
    ws_gantt = wb['Gantt Besoins']
    ws_cal_reel = wb['Calendrier réel']

    start_date = get_previous_monday()
    dates = [start_date + timedelta(weeks=i) for i in range(120)]
    for col, d in enumerate(dates, start=2):
        for ws in [ws_gantt, ws_cal_reel]:
            cell = ws.cell(4, col, d)
            cell.number_format = "dd/mm/yyyy"
            cell.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90)
            ws.column_dimensions[get_column_letter(col)].width = 5.0

    current_projects = []
    for row in range(3, ws_desc.max_row + 1):
        value = ws_desc.cell(row, 1).value
        if value and str(value).strip() and str(value).strip().lower() != "projet":
            current_projects.append(str(value).strip())
    st.session_state.projects = list(dict.fromkeys(current_projects))

    if not st.session_state.initial_rebuild_done:
        rebuild_gantt_sheet(ws_gantt, ws_desc, st.session_state.projects)
        rebuild_calendrier_sheet(ws_cal_reel, ws_desc, st.session_state.projects)
        st.session_state.initial_rebuild_done = True

    # 1. Ajouter un Projet
    st.subheader("1. Ajouter un Projet")
    new_proj = st.text_input("Nom du nouveau projet", key="new_proj")
    new_stat = st.selectbox("Statut initial", ["En soumission", "Contrat obtenu", "Abandonné"], key="new_stat")
    col1, col2 = st.columns(2)
    with col1:
        date_soumission = st.date_input("Date soumission", value=datetime(2025, 12, 12).date(), key="date_soumission")
    with col2:
        date_obtention = st.date_input("Date obtention", value=datetime(2026, 1, 16).date(), key="date_obtention")
    if st.button("Ajouter le projet") and new_proj:
        if any(p.lower() == new_proj.lower() for p in st.session_state.projects):
            st.error("❌ Un projet avec ce nom existe déjà !")
        else:
            last_project_row = 2
            for r in range(3, ws_desc.max_row + 1):
                if ws_desc.cell(r, 1).value and str(ws_desc.cell(r, 1).value).strip():
                    last_project_row = r
            next_row_desc = last_project_row + 1
            ws_desc.cell(next_row_desc, 1, new_proj)
            ws_desc.cell(next_row_desc, 2, new_stat)
            ws_desc.cell(next_row_desc, 3, date_soumission)
            ws_desc.cell(next_row_desc, 4, date_obtention)
            for c in range(5, 25):
                ws_desc.cell(next_row_desc, c, 0)
            st.session_state.projects.append(new_proj)
            rebuild_gantt_sheet(ws_gantt, ws_desc, st.session_state.projects)
            rebuild_calendrier_sheet(ws_cal_reel, ws_desc, st.session_state.projects)
            st.success(f"✅ {new_proj} ajouté !")
            st.rerun()

    # 2. Besoin projet approximatif
    st.subheader("2. Besoin projet approximatif")
    selected_approx = st.selectbox("Projet", st.session_state.projects, key="approx_select")
    row_approx = find_project_row(ws_desc, selected_approx)
    lit_approx = st.number_input("Besoin Lit approx", value=safe_float(ws_desc.cell(row_approx, 5).value) if row_approx else 0.0, min_value=0.0, key=f"lit_approx_{selected_approx}")
    dort_approx = st.number_input("Besoin dortoir approx (auto)", value=safe_float(ws_desc.cell(row_approx, 6).value) if row_approx else 0.0, min_value=0.0, disabled=True, key=f"dort_approx_{selected_approx}")
    bur_approx = st.number_input("Besoin module bureau approx", value=safe_float(ws_desc.cell(row_approx, 7).value) if row_approx else 0.0, min_value=0.0, key=f"bur_approx_{selected_approx}")
    vas_approx = st.number_input("Besoin module vaste approx", value=safe_float(ws_desc.cell(row_approx, 8).value) if row_approx else 0.0, min_value=0.0, key=f"vas_approx_{selected_approx}")
    if st.button("Enregistrer Besoin approximatif"):
        if row_approx:
            ws_desc.cell(row_approx, 5, lit_approx)
            dort_auto = math.ceil(lit_approx / 5.5) if lit_approx > 0 else 0
            ws_desc.cell(row_approx, 6, dort_auto)
            ws_desc.cell(row_approx, 7, bur_approx)
            ws_desc.cell(row_approx, 8, vas_approx)
            ws_desc.cell(row_approx, 9, math.floor(lit_approx * 0.3))
            ws_desc.cell(row_approx, 10, math.floor(dort_auto * 0.3))
            ws_desc.cell(row_approx, 11, math.floor(bur_approx * 0.3))
            ws_desc.cell(row_approx, 12, math.floor(vas_approx * 0.3))
            st.success("✅ Besoin approximatif + MAX NT + Dortoir auto enregistrés")
            st.rerun()

    # 3. Modifier infos projet
    st.subheader("3. Modifier infos projet existant")
    selected_edit = st.selectbox("Projet à modifier", st.session_state.projects, key="edit_select")
    row_edit = find_project_row(ws_desc, selected_edit)
    if row_edit:
        current_stat = str(ws_desc.cell(row_edit, 2).value or "En soumission")
        stat_options = ["En soumission", "Contrat obtenu", "Abandonné"]
        index_stat = stat_options.index(current_stat) if current_stat in stat_options else 0
        new_stat_edit = st.selectbox("Nouveau statut", stat_options, index=index_stat, key=f"stat_edit_{selected_edit}")
        current_date_sou = normalize_date(ws_desc.cell(row_edit, 3).value) or datetime(2025, 12, 12).date()
        new_date_soumission = st.date_input("Nouvelle date soumission", value=current_date_sou, key=f"date_soumission_edit_{selected_edit}")
        current_date_obt = normalize_date(ws_desc.cell(row_edit, 4).value) or datetime(2026, 1, 16).date()
        new_date_obtention = st.date_input("Nouvelle date obtention", value=current_date_obt, key=f"date_obtention_edit_{selected_edit}")
        if st.button("Enregistrer modification infos projet"):
            ws_desc.cell(row_edit, 2, new_stat_edit)
            ws_desc.cell(row_edit, 3, new_date_soumission)
            ws_desc.cell(row_edit, 4, new_date_obtention)
            rebuild_gantt_sheet(ws_gantt, ws_desc, st.session_state.projects)
            rebuild_calendrier_sheet(ws_cal_reel, ws_desc, st.session_state.projects)
            st.success(f"✅ Infos de {selected_edit} mises à jour")
            st.rerun()

    # 4. Capacité NT
    st.subheader("4. Capacité NT")
    if st.session_state.role in ["NT", "Admin"]:
        selected_cap = st.selectbox("Projet", st.session_state.projects, key="cap_select")
        row_cap = find_project_row(ws_desc, selected_cap)
        if row_cap:
            max_lit = safe_float(ws_desc.cell(row_cap, 9).value)
            max_bur = safe_float(ws_desc.cell(row_cap, 11).value)
            max_vas = safe_float(ws_desc.cell(row_cap, 12).value)
            if max_lit == 0: max_lit = math.floor(safe_float(ws_desc.cell(row_cap, 5).value) * 0.3)
            if max_bur == 0: max_bur = math.floor(safe_float(ws_desc.cell(row_cap, 7).value) * 0.3)
            if max_vas == 0: max_vas = math.floor(safe_float(ws_desc.cell(row_cap, 8).value) * 0.3)
            st.info(f"**MAX NT autorisé** : Lit {max_lit} | Bureau {max_bur} | Vaste {max_vas}")
            cap_deja_saisi = row_cap and any(safe_float(ws_desc.cell(row_cap, c).value) > 0 for c in [13,15,16])
            cap_lit = st.number_input("Capacité NT Lit", value=safe_float(ws_desc.cell(row_cap, 13).value) if row_cap else 0.0, min_value=0.0, disabled=cap_deja_saisi, key=f"cap_lit_{selected_cap}")
            cap_bur = st.number_input("Capacité NT Bureau", value=safe_float(ws_desc.cell(row_cap, 15).value) if row_cap else 0.0, min_value=0.0, disabled=cap_deja_saisi, key=f"cap_bur_{selected_cap}")
            cap_vas = st.number_input("Capacité NT Vaste", value=safe_float(ws_desc.cell(row_cap, 16).value) if row_cap else 0.0, min_value=0.0, disabled=cap_deja_saisi, key=f"cap_vas_{selected_cap}")
            dort_cap = math.ceil(cap_lit / 5.5) if cap_lit > 0 else 0
            st.info(f"**Dortoir NT calculé automatiquement** : {dort_cap}")
            if st.button("Enregistrer Capacité NT") and not cap_deja_saisi and row_cap:
                ws_desc.cell(row_cap, 13, cap_lit)
                ws_desc.cell(row_cap, 14, dort_cap)
                ws_desc.cell(row_cap, 15, cap_bur)
                ws_desc.cell(row_cap, 16, cap_vas)
                ws_desc.cell(row_cap, 17, max(0, safe_float(ws_desc.cell(row_cap, 5).value) - cap_lit))
                ws_desc.cell(row_cap, 18, max(0, safe_float(ws_desc.cell(row_cap, 6).value) - dort_cap))
                ws_desc.cell(row_cap, 19, max(0, safe_float(ws_desc.cell(row_cap, 7).value) - cap_bur))
                ws_desc.cell(row_cap, 20, max(0, safe_float(ws_desc.cell(row_cap, 8).value) - cap_vas))
                st.success("✅ Capacité NT + Dortoir auto + Besoin à combler enregistrés")
                st.rerun()
    else:
        st.warning("❌ Seuls les utilisateurs NT peuvent accéder à la Capacité NT")

    # 5. Engagement RL
    st.subheader("5. Engagement RL")
    if st.session_state.role in ["RL", "Admin"]:
        selected_eng = st.selectbox("Projet", st.session_state.projects, key="eng_select")
        row_eng = find_project_row(ws_desc, selected_eng)
        eng_deja_saisi = row_eng and any(safe_float(ws_desc.cell(row_eng, c).value) > 0 for c in range(21, 25))
        lit_eng = st.number_input("Besoin Lit (Engagement)", value=safe_float(ws_desc.cell(row_eng, 21).value) if row_eng else 0.0, min_value=0.0, disabled=eng_deja_saisi, key=f"lit_eng_{selected_eng}")
        bur_eng = st.number_input("Besoin module bureau (Engagement)", value=safe_float(ws_desc.cell(row_eng, 23).value) if row_eng else 0.0, min_value=0.0, disabled=eng_deja_saisi, key=f"bur_eng_{selected_eng}")
        vas_eng = st.number_input("Besoin module vaste (Engagement)", value=safe_float(ws_desc.cell(row_eng, 24).value) if row_eng else 0.0, min_value=0.0, disabled=eng_deja_saisi, key=f"vas_eng_{selected_eng}")
        if eng_deja_saisi:
            st.warning("Engagement RL déjà saisi → modification bloquée.")
        if st.button("Enregistrer Engagement RL") and not eng_deja_saisi and row_eng:
            ws_desc.cell(row_eng, 21, lit_eng)
            ws_desc.cell(row_eng, 22, math.ceil(lit_eng / 5.5) if lit_eng else 0)
            ws_desc.cell(row_eng, 23, bur_eng)
            ws_desc.cell(row_eng, 24, vas_eng)
            st.success(f"Engagement RL enregistré pour {selected_eng}")
            st.rerun()
    else:
        st.warning("❌ Seuls les utilisateurs RL peuvent accéder à l'Engagement RL")

    # 6. Saisie par période
    st.subheader("6. Saisie par période (Gantt Besoins)")
    valid_period_projects = [p for p in st.session_state.projects if get_project_status(ws_desc, p) != "Abandonné"]
    selected_period = st.selectbox("Projet", valid_period_projects, key="period")
    col1, col2 = st.columns(2)
    with col1:
        date_debut = st.date_input("De", value=datetime(2026, 3, 1).date(), key="date_debut_gantt")
    with col2:
        date_fin = st.date_input("À", value=datetime(2026, 11, 30).date(), key="date_fin_gantt")
    lit = st.number_input("Besoin Lit", min_value=0.0, value=100.0, key="gantt_lit")
    bureau = st.number_input("Besoin module bureau", min_value=0.0, value=2.0, key="gantt_bureau")
    vaste = st.number_input("Besoin module vaste", min_value=0.0, value=3.0, key="gantt_vaste")
    if st.button("Appliquer période au Gantt Besoins"):
        proj_row = None
        for r in range(5, ws_gantt.max_row + 1, 6):
            cell_name = str(ws_gantt.cell(r, 1).value or "").strip()
            if cell_name.split(" - ")[0] == selected_period:
                proj_row = r
                break
        if proj_row:
            dortoir = math.ceil(lit / 5.5) if lit > 0 else 0
            for col in range(2, ws_gantt.max_column + 1):
                d = dates[col-2] if col-2 < len(dates) else None
                if d and date_debut <= d <= date_fin:
                    ws_gantt.cell(proj_row + 1, col, lit)
                    ws_gantt.cell(proj_row + 2, col, dortoir)
                    ws_gantt.cell(proj_row + 3, col, bureau)
                    ws_gantt.cell(proj_row + 4, col, vaste)
            st.success(f"Période appliquée pour {selected_period}")
            st.rerun()

    # VÉRIFICATION GAPS
    st.subheader("Vérification des semaines vides dans Gantt Besoins")
    if st.button("🔍 Vérifier les gaps dans Gantt"):
        gaps = check_gantt_gaps(ws_gantt)
        if gaps:
            st.error("**Semaines vides détectées :**")
            for g in gaps:
                st.markdown(f"- {g}")
            st.session_state.gantt_gap_confirmed = False
        else:
            st.success("✅ Aucun gap détecté")
            st.session_state.gantt_gap_confirmed = True
    confirm_gap = st.checkbox("Je confirme que les semaines vides sont correctes (report de projet, etc.)",
                              value=st.session_state.gantt_gap_confirmed, key="confirm_gap")
    st.session_state.gantt_gap_confirmed = confirm_gap

    # 7. Saisie Calendrier Réel
    st.subheader("7. Saisie Calendrier Réel (avec report automatique)")
    obtained_projects = [p for p in st.session_state.projects if get_project_status(ws_desc, p) == "Contrat obtenu"]
    if not obtained_projects:
        st.warning("Aucun projet avec statut **Contrat obtenu** pour l'instant.")
    else:
        selected_real = st.selectbox("Projet", obtained_projects, key="real")
        selected_date = st.date_input("Date de la semaine", value=get_previous_monday(), key="week_input")
        week_monday = get_monday_of_week(selected_date)
        week_key = week_monday.strftime("%Y%m%d")
        st.info(f"**Semaine enregistrée (lundi) : {week_monday.strftime('%d/%m/%Y')}**")
        proj_row = find_project_row(ws_cal_reel, selected_real, start_row=5)
        if not proj_row:
            st.error("Bloc projet non trouvé.")
        else:
            col = find_or_create_week_column(ws_cal_reel, week_monday)
            existing_data = {
                "dortoir_rl": safe_float(ws_cal_reel.cell(proj_row + 1, col).value),
                "bureau_rl": safe_float(ws_cal_reel.cell(proj_row + 2, col).value),
                "vaste_rl": safe_float(ws_cal_reel.cell(proj_row + 3, col).value),
                "dortoir_nt": safe_float(ws_cal_reel.cell(proj_row + 5, col).value),
                "bureau_nt": safe_float(ws_cal_reel.cell(proj_row + 6, col).value),
                "vaste_nt": safe_float(ws_cal_reel.cell(proj_row + 7, col).value)
            }
            if all(v == 0 for v in existing_data.values()):
                prev_col = find_last_filled_column(ws_cal_reel, proj_row, col)
                if prev_col:
                    st.info("🔄 Valeurs reportées automatiquement de la dernière saisie")
                    existing_data = {
                        "dortoir_rl": safe_float(ws_cal_reel.cell(proj_row + 1, prev_col).value),
                        "bureau_rl": safe_float(ws_cal_reel.cell(proj_row + 2, prev_col).value),
                        "vaste_rl": safe_float(ws_cal_reel.cell(proj_row + 3, prev_col).value),
                        "dortoir_nt": safe_float(ws_cal_reel.cell(proj_row + 5, prev_col).value),
                        "bureau_nt": safe_float(ws_cal_reel.cell(proj_row + 6, prev_col).value),
                        "vaste_nt": safe_float(ws_cal_reel.cell(proj_row + 7, prev_col).value)
                    }
            dortoir_rl = st.number_input("Dortoir RL", min_value=0, value=int(existing_data["dortoir_rl"]), key=f"dortoir_rl_{selected_real}_{week_key}")
            bureau_rl = st.number_input("Bureau RL", min_value=0, value=int(existing_data["bureau_rl"]), key=f"bureau_rl_{selected_real}_{week_key}")
            vaste_rl = st.number_input("Vaste RL", min_value=0, value=int(existing_data["vaste_rl"]), key=f"vaste_rl_{selected_real}_{week_key}")
            dortoir_nt = st.number_input("Dortoir NT", min_value=0, value=int(existing_data["dortoir_nt"]), key=f"dortoir_nt_{selected_real}_{week_key}")
            bureau_nt = st.number_input("Bureau NT", min_value=0, value=int(existing_data["bureau_nt"]), key=f"bureau_nt_{selected_real}_{week_key}")
            vaste_nt = st.number_input("Vaste NT", min_value=0, value=int(existing_data["vaste_nt"]), key=f"vaste_nt_{selected_real}_{week_key}")
            if st.button("Enregistrer saisie réelle pour cette semaine", type="primary"):
                last_col = find_last_filled_column(ws_cal_reel, proj_row, col)
                if last_col and last_col < col - 1:
                    backfill_intermediate_weeks(ws_cal_reel, proj_row, last_col, col)
                    st.info(f"🔄 {col - last_col - 1} semaine(s) intermédiaire(s) remplie(s) automatiquement")
                total_rl = dortoir_rl + bureau_rl + vaste_rl
                total_nt = dortoir_nt + bureau_nt + vaste_nt
                total_module = total_rl + total_nt
                ws_cal_reel.cell(proj_row + 1, col, dortoir_rl)
                ws_cal_reel.cell(proj_row + 2, col, bureau_rl)
                ws_cal_reel.cell(proj_row + 3, col, vaste_rl)
                ws_cal_reel.cell(proj_row + 4, col, total_rl)
                ws_cal_reel.cell(proj_row + 5, col, dortoir_nt)
                ws_cal_reel.cell(proj_row + 6, col, bureau_nt)
                ws_cal_reel.cell(proj_row + 7, col, vaste_nt)
                ws_cal_reel.cell(proj_row + 8, col, total_nt)
                ws_cal_reel.cell(proj_row + 9, col, total_module)
                st.success(f"✅ Saisie enregistrée pour {selected_real} – Semaine du {week_monday.strftime('%d/%m/%Y')}")
                st.rerun()

    # 8. Rattrapage
    st.subheader("8. Rattrapage (Cumul Global identique pour tous les projets)")
    st.info("✅ Pic_Réel + Max_RL + Max_NT + % RL/NT + Rattrapage_Créé/Utilisé + **Cumul_Déficit_Global**")
    if st.button("🔄 Mettre à jour Rattrapage maintenant"):
        update_rattrapage_sheet(wb)
        st.success("✅ Onglet Rattrapage mis à jour avec Cumul Global !")
        st.rerun()

    # EXPORT
    if st.button("Exporter Maj", type="primary"):
        if not st.session_state.gantt_gap_confirmed:
            st.error("❌ Vous devez cocher la case de confirmation des semaines vides avant d'exporter !")
        else:
            with st.spinner("Export en cours..."):
                update_rattrapage_sheet(wb)
                apply_all_styling(wb)
                timestamp = datetime.now(ZoneInfo("America/Montreal")).strftime("%Y-%m-%d_%H-%M")
                output_file = f'besoins_maj_{timestamp}.xlsx'
                wb.save(output_file)
                with open(output_file, 'rb') as f:
                    st.download_button(
                        label="📥 Télécharger le fichier mis à jour",
                        data=f,
                        file_name=output_file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                send_email(ADMIN_EMAIL, "Export mis à jour", "Voici la dernière version du fichier.", output_file)
                send_to_all_users("Export mis à jour - Gestion Contrats RL/NT", "Voici la dernière version du fichier.", output_file)
            st.success("✅ Export terminé + envoyé par email à tous les utilisateurs !")

else:
    st.warning("Upload ton fichier **Modèle Base.xlsx** pour commencer.")

st.caption("✅ CODE 100% COMPLET – Sous-Totaux + TOTAL restaurés + lien dans le mail + tout le reste intact")
