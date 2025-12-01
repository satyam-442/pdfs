import pandas as pd
import win32com.client as win32

# ---------------- CONFIG ---------------- #
EXCEL_PATH = "PRB_data.xlsx"
SHEETS = ["BreachedPRB", "AboutToBreachPRB"]
OPEN_EMAIL_IN_DRAFT = True
TEST_TO_EMAIL = "shivramtripathi700@gmail.com"
# ---------------------------------------- #

# Final display columns (Opened excluded)
BREACHED_COLUMNS = [
    "Number", "Assignment Group", "Priority", "Breach Type", "Service",
    "Assigned to", "BO director", "Manager", "Short description",
    "Age", "Related Inc#", "Last Updated"
]

ABOUT_COLUMNS = [
    "Number", "Assignment Group", "Priority", "Breach Type", "Service",
    "Assigned to", "BO director", "Manager", "Short description",
    "Age", "Time to breach", "Related Inc#", "Last Updated"
]


# -------------- HTML STYLE -------------- #
HTML_STYLE = """
<style>
    body {
        font-family: Arial, sans-serif;
        font-size: 12px;
        color: #111;
    }
    h2 {
        margin-top: 20px;
        margin-bottom: 5px;
        font-size: 14px;
    }
    table {
        border-collapse: collapse;
        width: 100%;
        margin-bottom: 20px;
    }
    th {
        background-color: #f2f2f2;
        padding: 6px;
        border: 1px solid #ccc;
        font-weight: bold;
        white-space: nowrap;
    }
    td {
        padding: 6px;
        border: 1px solid #ccc;
        vertical-align: top;
    }
    .highlight {
        background-color: #fff59d; /* yellow */
    }
</style>
"""


# -------------- LOAD DATA -------------- #
dfs = []
for sheet in SHEETS:
    print(f"Loading sheet: {sheet}")
    df = pd.read_excel(EXCEL_PATH, sheet_name=sheet, dtype=str)
    df.fillna("", inplace=True)

    df["AssigneeEmail"] = ""   
    df["ManagerEmail"] = ""
    df["__Sheet"] = sheet

    dfs.append(df)

combined = pd.concat(dfs, ignore_index=True)
print("Combined rows:", len(combined))


# -------------- HELPER -------------- #
def escape_html(text):
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def build_html_table(df, columns):
    """Generate HTML table for dataframe."""
    html = "<table><tr>"
    for col in columns:
        html += f"<th>{escape_html(col)}</th>"
    html += "</tr>"

    for _, row in df.iterrows():
        assigned = str(row.get("Assigned to", "")).strip()
        html += "<tr>"

        for col in columns:
            val = escape_html(row.get(col, ""))

            if col == "Number" and assigned == "":
                html += f"<td class='highlight'>{val}</td>"
            else:
                html += f"<td>{val}</td>"

        html += "</tr>"

    html += "</table>"
    return html


# -------------- EMAIL SENDER -------------- #
def send_email(assignment_group, html_body):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.Subject = f"EDO PRB - {assignment_group}"
    mail.HTMLBody = html_body
    mail.To = TEST_TO_EMAIL     # SAFE FOR TESTING

    if OPEN_EMAIL_IN_DRAFT:
        mail.Display()
    else:
        mail.Send()


# -------------- MAIN LOOP -------------- #
print("\n--- Generating HTML Emails per Assignment Group ---\n")

assignment_groups = combined["Assignment Group"].unique()

for ag in assignment_groups:

    if str(ag).strip() == "":
        continue

    html_body = HTML_STYLE
    html_body += f"<h2>Assignment Group: {escape_html(ag)}</h2>"

    df_group = combined[combined["Assignment Group"] == ag]

    breached_rows = df_group[df_group["__Sheet"] == "BreachedPRB"]
    about_rows = df_group[df_group["__Sheet"] == "AboutToBreachPRB"]

    # Breached table
    if not breached_rows.empty:
        html_body += f"<h2>EDO Open PRBs in Breach (P3–P5) – {escape_html(ag)}</h2>"
        html_body += build_html_table(breached_rows, BREACHED_COLUMNS + ["AssigneeEmail", "ManagerEmail"])

    # About-to-breach table
    if not about_rows.empty:
        html_body += f"<h2>EDO Open PRBs about to Breach (P3–P5) – {escape_html(ag)}</h2>"
        html_body += build_html_table(about_rows, ABOUT_COLUMNS + ["AssigneeEmail", "ManagerEmail"])

    # No rows → skip email
    if breached_rows.empty and about_rows.empty:
        continue

    print(f"Sending HTML email for Assignment Group: {ag}")
    send_email(ag, html_body)

print("\n--- HTML Emails Generated Successfully (Testing Mode) ---\n")
