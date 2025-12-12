import gspread
from google.oauth2.service_account import Credentials
from flask import Flask, request, session, render_template_string, redirect, url_for
from datetime import datetime
import sqlite3
import sys, os

# ----------------------------------------------------------
# FONCTION COMPATIBLE EXE POUR CHARGER LES FICHIERS
# ----------------------------------------------------------
def resource_path(relative_path):
    """Obtenir le vrai chemin d’un fichier, compatible PyInstaller"""
    try:
        base_path = sys._MEIPASS   # dossier temporaire PyInstaller
    except Exception:
        base_path = os.path.abspath(".")  # dossier normal en mode .py
    return os.path.join(base_path, relative_path)

# ----------------------------------------------------------
# FLASK
# ----------------------------------------------------------
app = Flask(__name__)
app.secret_key = "secret_key_ici"

# ----------------------------------------------------------
# GOOGLE SHEETS - CHARGEMENT DU JSON
# ----------------------------------------------------------
scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

credentials_file = resource_path("googleapi.json")

creds = Credentials.from_service_account_file(credentials_file, scopes=scope)
client = gspread.authorize(creds)

googlsheet = client.open("PROJECT")
authentification = googlsheet.worksheet("AUTH")
feuilleauth = googlsheet.worksheet("AGENT CONFIG")
dispatch = googlsheet.worksheet("dispatch")
resultat = googlsheet.worksheet("DEMANDeandresult")
contenusheet = authentification.acell("A1").value
nouveauTraitement = ""


# ----------------------------------------------------------
# ROUTE LOGIN
# ----------------------------------------------------------
@app.route('/', methods=["GET", "POST"])
def import_route():
    if request.method == 'POST':

        nom = request.form.get("nom")
        mdp = request.form.get("mdp")

        data = feuilleauth.get("A:C")
        datadispatch = dispatch.get("A:D")

        filtered_data = [(row[0], row[2]) for row in data if row and row[0] == nom and row[1] == mdp]
        assignationcsa = [(row[2]) for row in data if row and row[0] == nom]
        dispatch_rows = [(row[0], row[1], row[2], row[3]) for row in datadispatch if row and row[0] == nom]

        if not filtered_data:
            return render_template_string(contenusheet)

        # --- SQLite ---
        conn = sqlite3.connect("ma_base.db")
        c = conn.cursor()

        c.execute("DROP TABLE IF EXISTS ma_table_base_assignation")
        c.execute("DROP TABLE IF EXISTS ma_table_base_a_traite")
        c.execute("DROP TABLE IF EXISTS ma_table_base_resultat")

        c.execute("""
            CREATE TABLE IF NOT EXISTS ma_table_base_assignation(
                assignation TEXT
            )
        """)

        c.execute("""
            CREATE TABLE IF NOT EXISTS ma_table_base_a_traite (
                col1 TEXT,
                col2 TEXT,
                col3 TEXT,
                id TEXT
            )
        """)

        c.execute("""
            CREATE TABLE IF NOT EXISTS ma_table_base_resultat (
                id TEXT
            )
        """)

        c.executemany("INSERT INTO ma_table_base_a_traite (col1, col2, col3, id) VALUES (?, ?, ?, ?)", dispatch_rows)
        c.execute("INSERT INTO ma_table_base_assignation (assignation) VALUES (?)", assignationcsa)
        conn.commit()
        conn.close()

        print(dispatch_rows)
        return redirect(url_for('traitement'))

    return render_template_string(contenusheet)


# ----------------------------------------------------------
# ROUTE TRAITEMENT
# ----------------------------------------------------------
@app.route('/Traitement', methods=["GET", "POST"])
def traitement():

    if request.method == "POST":

        agent_post = request.form.get("nom")
        action = request.form.get("action")
        dateHeure = request.form.get("dateHeure")
        allinput = request.form.get("allinput")
        qualification = request.form.get("qualification")
        pausetotal = request.form.get("pausetotal")

        # ---------------------------------------------------
        # TRAITEMENT
        # ---------------------------------------------------
        if action == "traitement":
            nouveauTraitement = request.form.get("nouveauTraitement")
            id = request.form.get("id")
            campagne = request.form.get("campagne")

            statut2 = request.form.get("statut2")

            resultat.append_row([
                agent_post, dateHeure, datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                allinput, nouveauTraitement, statut2, campagne
            ])

            conn = sqlite3.connect("ma_base.db")
            c = conn.cursor()

            c.execute(""" INSERT INTO ma_table_base_resultat (id) VALUES (?) """, (id,))
            conn.commit()

            c.execute("""
                SELECT * 
                FROM ma_table_base_a_traite AS t1
                LEFT JOIN ma_table_base_resultat AS t2 ON t1.id = t2.id
                JOIN ma_table_base_assignation AS t3 ON t1.col3 = t3.assignation
                WHERE t2.id IS NULL
                LIMIT 1;
            """)

            row = c.fetchone()

            if row is not None:
                conn.close()

                nouveauTraitement = "TRAITEMENT"
                donneTraitement = row[1]
                statut2 = "TRAITEMENT"
                print(row)
                id = row[3]
                agent = row[0]
                campagne = row[2]
                campagneparcouru = googlsheet.worksheet(campagne)
                celluleHTML = campagneparcouru.acell("A1").value

                print(id)

                return render_template_string(
                    celluleHTML,
                    agent=agent_post,
                    donneTraitement=donneTraitement,
                    nouveauTraitement=nouveauTraitement,
                    statut2=statut2,
                    id=id,
                    campagne=campagne
                )

            else:
                c.execute("SELECT col1 FROM ma_table_base_a_traite")
                row = c.fetchone()
                agent = row[0]
                statut2 = "ATTENTE"

                c.execute("SELECT assignation FROM ma_table_base_assignation")
                campagnes = c.fetchone()
                campagne = campagnes[0]

                nouveauTraitement = "ATTENTE"
                campagneparcouru = googlsheet.worksheet(campagne)
                celluleHTML = campagneparcouru.acell("A1").value

                return render_template_string(
                    celluleHTML,
                    agent=agent,
                    nouveauTraitement=nouveauTraitement,
                    statut2=statut2,
                    campagne=campagne
                )

        # ---------------------------------------------------
        # PAUSE
        # ---------------------------------------------------
        elif action == "pause":
            statut2 = request.form.get("statut2")
            pausetotal = request.form.get("pausetotal")
            campagne = request.form.get("campagne")
            id = request.form.get("id")
            allinput = request.form.get("allinput")
            nouveauTraitement = request.form.get("nouveauTraitement")

            resultat.append_row([
                agent_post, dateHeure, datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                allinput, nouveauTraitement, statut2, campagne
            ])

            nouveauTraitement = "PAUSE"
            statut2 = pausetotal

            conn = sqlite3.connect("ma_base.db")
            c = conn.cursor()

            c.execute(""" INSERT INTO ma_table_base_resultat (id) VALUES (?) """, (id,))
            c.execute("SELECT assignation FROM ma_table_base_assignation")
            conn.commit()

            row = c.fetchone()
            campagne = row[0]
            campagneparcouru = googlsheet.worksheet(campagne)
            celluleHTML = campagneparcouru.acell("A1").value

            return render_template_string(
                celluleHTML,
                agent=agent_post,
                nouveauTraitement=nouveauTraitement,
                statut2=statut2,
                campagne=campagne
            )

        # ---------------------------------------------------
        # LOGOUT
        # ---------------------------------------------------
        elif action == "logout":
            nouveauTraitement = request.form.get("nouveauTraitement")
            statut2 = request.form.get("statut2")
            allinput = request.form.get("allinput")
            campagne = request.form.get("campagne")

            resultat.append_row([
                agent_post, dateHeure, datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                allinput, nouveauTraitement, statut2, campagne
            ])

            conn = sqlite3.connect("ma_base.db")
            c = conn.cursor()
            c.execute("DROP TABLE IF EXISTS ma_table_base_a_traite")
            c.execute("DROP TABLE IF EXISTS ma_table_base_resultat")
            c.execute("DROP TABLE IF EXISTS ma_table_base_assignation")
            conn.commit()
            conn.close()

            return redirect(url_for('import_route'))

    else:

        action = ""
        allinput = "-"
        statut2 = "TRAITEMENT"
        campagne = request.form.get("campagne")
        pausetotal = request.form.get("pausetotal")

        conn = sqlite3.connect("ma_base.db")
        c = conn.cursor()

        c.execute("""
            SELECT * 
            FROM ma_table_base_a_traite AS t1
            LEFT JOIN ma_table_base_resultat AS t2 ON t1.id = t2.id
            JOIN ma_table_base_assignation AS t3 ON t1.col3 = t3.assignation
            WHERE t2.id IS NULL
            LIMIT 1;
        """)

        row = c.fetchone()

        if row is not None:

            allinput = "-"
            nouveauTraitement = "TRAITEMENT"
            id = row[3]
            agent = row[0]
            campagne = row[2]
            donneTraitement = row[1]
            campagneparcouru = googlsheet.worksheet(campagne)
            celluleHTML = campagneparcouru.acell("A1").value

            resultat.append_row([
                agent, "", datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                allinput, "ATTENTE", statut2, campagne
            ])

            return render_template_string(
                celluleHTML,
                agent=agent,
                donneTraitement=donneTraitement,
                nouveauTraitement=nouveauTraitement,
                statut2=statut2,
                id=id,
                campagne=campagne
            )

        else:
            c.execute("SELECT col1 FROM ma_table_base_a_traite")
            row = c.fetchone()
            agent = row[0]
            statut2 = "Attente"

            c.execute("SELECT assignation FROM ma_table_base_assignation")
            campagnes = c.fetchone()
            campagne = campagnes[0]

            nouveauTraitement = "Attente"
            campagneparcouru = googlsheet.worksheet(campagne)
            celluleHTML = campagneparcouru.acell("A1").value

            resultat.append_row([
                agent, "", datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                allinput, "ATTENTE", statut2, campagne
            ])

            return render_template_string(
                celluleHTML,
                agent=agent,
                nouveauTraitement=nouveauTraitement,
                statut2=statut2,
                campagne=campagne
            )


# ----------------------------------------------------------
# RUN
# ----------------------------------------------------------
if __name__ == '__main__':
    app.run(debug=True)
