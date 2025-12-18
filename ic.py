import pandas as pd
import win32com.client as win32

# ================= CONFIG ================= #
EXCEL_PATH = "PRB_data.xlsx"
SHEETS = ["BreachedPRB", "AboutToBreachPRB"]

OPEN_EMAIL_IN_DRAFT = True   # True = save to Drafts, False = Send

CC_LIST = [
    "shivramtripathi700@gmail.com",
    "qa_manager@example.com"
]

BASE_PRB_URL = (
    "https://capitalgroup.service-now.com/"
    "nav_to.do?uri=problem.do?sysparm_query=number="
)
# ========================================= #

# -------- COLUMN DEFINITIONS -------- #
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

COLUMN_WIDTHS = {
    "Number": 120,
    "Assignment Group": 220,
    "Priority": 90,
    "Breach Type": 120,
    "Service": 120,
    "Assigned to": 140,
    "BO director": 140,
    "Manager": 140,
    "Short description": 320,
    "Age": 60,
    "Time to breach": 110,
    "Related Inc#": 120,
    "Last Updated": 120
}
# ----------------------------------- #

# ========== HTML STYLE ==========
HTML_STYLE = """
<style>
body { font-family: Arial, sans-serif; font-size: 12px; color: #111; }
h2 { font-size: 14px; margin: 16px 0 6px; }
table { border-collapse: collapse; margin-bottom: 16px; }
th, td { border: 1px solid #cfcfcf; padding: 6px; vertical-align: top; }
th { background-color: #f2f2f2; font-weight: bold; white-space: nowrap; }
.highlight { background-color: #fff59d; }
a { color: #0563c1; text-decoration: underline; }
</style>
"""

# ========== LOAD DATA ==========
dfs = []

for sheet in SHEETS:
    df = pd.read_excel(EXCEL_PATH, sheet_name=sheet, dtype=str)
    df.fillna("", inplace=True)

    # Email columns (safe placeholders / real emails in prod)
    df["AssigneeEmail"] = df["Assigned to"]
    df["ManagerEmail"] = df["Manager"]

    df["__Sheet"] = sheet
    dfs.append(df)

combined = pd.concat(dfs, ignore_index=True)

# ========== HELPERS ==========
def escape_html(text):
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def build_html_table(df, columns):
    html = "<table cellpadding='4' cellspacing='0' border='1'>"

    # Header
    html += "<tr>"
    for col in columns:
        width = COLUMN_WIDTHS.get(col, 120)
        html += f"<th width='{width}' align='left'>{escape_html(col)}</th>"
    html += "</tr>"

    # Rows
    for _, row in df.iterrows():
        assigned = row.get("Assigned to", "").strip()
        html += "<tr>"

        for col in columns:
            width = COLUMN_WIDTHS.get(col, 120)
            val = escape_html(row.get(col, ""))

            # Hyperlink PRB Number
            if col == "Number" and val:
                link = f"{BASE_PRB_URL}{val}"
                cell_value = f"<a href='{link}'>{val}</a>"
            else:
                cell_value = val

            # Highlight Number if Assigned To empty
            if col == "Number" and assigned == "":
                html += f"<td width='{width}' class='highlight'>{cell_value}</td>"
            else:
                html += f"<td width='{width}'>{cell_value}</td>"

        html += "</tr>"

    html += "</table>"
    return html


def determine_to_cc(df_group):
    TO = set()
    CC = set(CC_LIST)

    for _, row in df_group.iterrows():
        assignee = row.get("AssigneeEmail", "").strip()
        manager = row.get("ManagerEmail", "").strip()

        if assignee == "":
            if manager:
                TO.add(manager)
        else:
            TO.add(assignee)
            if manager:
                CC.add(manager)

    return list(TO), list(CC)


def send_email(subject, html_body, to_list, cc_list):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.Subject = subject
    mail.HTMLBody = html_body
    mail.To = "; ".join(to_list)
    mail.CC = "; ".join(cc_list)

    if OPEN_EMAIL_IN_DRAFT:
        mail.Save()   # silent draft
    else:
        mail.Send()


# ========== MAIN PROCESS ==========
for ag in combined["Assignment Group"].unique():

    if not ag.strip():
        continue

    df_group = combined[combined["Assignment Group"] == ag]

    breached_rows = df_group[df_group["__Sheet"] == "BreachedPRB"]
    about_rows = df_group[df_group["__Sheet"] == "AboutToBreachPRB"]

    if breached_rows.empty and about_rows.empty:
        continue

    html_body = HTML_STYLE
    html_body += f"""
    Dear All,<br><br>
    Your action is needed as there are P3–P5 Problem Records (PRB) assigned to you
    that are about to breach or have already breached their Service Level Agreements (SLA).
    <br><br>
    """

    if not breached_rows.empty:
        html_body += f"<h2>EDO Open PRBs in Breach (P3–P5) – {escape_html(ag)}</h2>"
        html_body += build_html_table(breached_rows, BREACHED_COLUMNS)

    if not about_rows.empty:
        html_body += f"<h2>EDO Open PRBs about to Breach (P3–P5) – {escape_html(ag)}</h2>"
        html_body += build_html_table(about_rows, ABOUT_COLUMNS)

    html_body += "<br>Regards,<br>4DATA Service Desk"

    TO_LIST, CC_LIST_FINAL = determine_to_cc(df_group)

    if not TO_LIST:
        continue

    send_email(
        subject=f"EDO PRB - {ag}",
        html_body=html_body,
        to_list=TO_LIST,
        cc_list=CC_LIST_FINAL
    )

print("✅ PRB HTML emails created successfully (Draft mode).")
