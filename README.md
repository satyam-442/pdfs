<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Show Code Exact</title>
    <style>
        pre {
            background: #1e1e1e;
            color: #ccc;
            padding: 15px;
            font-size: 14px;
            overflow-x: auto;
            white-space: pre;
            border-radius: 6px;
        }
    </style>
</head>
<body>

<pre id="codeBlock"></pre>

<script>
const code = String.raw`
# pdfs
import pandas as pd
import win32com.client as win32

# ---------------- CONFIG ---------------- #
EXCEL_PATH = "PRB_data.xlsx"   # <-- change to your PRB file path
DOMAIN = "example.com"         # fallback domain for name->email conversion
CC_LIST = ["shivramtripathi700@gmail.com", "qa_manager@example.com"]

OPEN_EMAIL_IN_DRAFT = True     # True => mail.Display(); False => mail.Send()
SHEETS = ["BreachedPRB", "AboutToBreachPRB"]

# Final column lists (as you specified)
BREACHED_COLUMNS = ["Number", "Assignment Group", "Priority", "Breach Type", "Service", "Assigned to", "BO director", "Manager", "Short description", "Age", "Related Inc#", "Last Updated"]

ABOUT_COLUMNS = ["Number", "Assignment Group", "Priority", "Breach Type", "Service", "Assigned to", "BO director", "Manager", "Short description", "Age", "Time to breach", "Related Inc#", "Last Updated"]

def to_email(value):
    """Convert a name or email-like string to an email address (simple heuristic)."""
    v = str(value).strip()
    if not v:
        return ""
    if "@" in v:
        return v
    parts = v.lower().replace(",", "").split()
    if len(parts) >= 2:
        return f"{parts[0]}.{parts[-1]}@DOMAIN"
    return f"{parts[0]}@DOMAIN"

# Load sheets and tag with sheet name
dfs = []
for s in SHEETS:
    df = pd.read_excel(EXCEL_PATH, sheet_name=s, dtype=str)
    df.fillna("", inplace=True)    # convert NaN -> ""
    df["__Sheet"] = s
    dfs.append(df)

combined = pd.concat(dfs, ignore_index=True)

# Build HTML for a manager
def build_html_for_manager(manager_name: str) -> str:
    mgr = str(manager_name).strip()
    df_mgr = combined[combined["Manager"].str.strip() == mgr]
    if df_mgr.empty:
        return ""
    css = \"\"\"
    <style>
      body { font-family: Arial, sans-serif; font-size:11px; color:#111; }
      .header { margin-bottom:10px; }
      .group-title { font-weight:700; margin:12px 0 6px 0; }
      table { border-collapse: collapse; width:100%; margin-bottom:12px; }
      th, td { border: 1px solid #cfcfcf; padding:6px 8px; font-size:11px; vertical-align:top; }
      th { background:#f3f3f3; white-space:nowrap; font-weight:600; }
      .highlight { background-color: #fff59d; }
      .small { font-size:10px; color:#555; margin-bottom:8px; }
      .sheet-title { margin-top:8px; margin-bottom:6px; font-weight:600; }
    </style>
    \"\"\"
    html = [f"<html><head>{css}</head><body>"]
    html.append(f"<div class='header'><b>Manager:</b> {mgr}</div>")
    for ag, _ in df_mgr.groupby("Assignment Group", sort=False):
        ag_display = ag if str(ag).strip() else "(No Assignment Group)"
        df_group = df_mgr[df_mgr["Assignment Group"] == ag]

        df_breached = df_group[df_group["__Sheet"] == "BreachedPRB"]
        if not df_breached.empty:
            title = f"EDO Open PRBs in Breach (P3–P5) – {ag_display}"
            html.append(f"<div class='group-title'>{title}</div>")

            cols = [c for c in BREACHED_COLUMNS if c in df_breached.columns]
            html.append("<table>")
            html.append("<tr>" + "".join(f"<th>{c}</th>" for c in cols) + "</tr>")

            for _, row in df_breached.iterrows():
                assigned = str(row.get("Assigned to", "")).strip()
                row_cells = []
                for col in cols:
                    val = "" if pd.isna(row.get(col, "")) else str(row.get(col, ""))
                    if col == "Number" and assigned == "":
                        cell = f"<td class='highlight'>{escape_html(val)}</td>"
                    else:
                        cell = f"<td>{escape_html(val)}</td>"
                    row_cells.append(cell)
                html.append("<tr>" + "".join(row_cells) + "</tr>")
            html.append("</table>")
            
`;

document.getElementById("codeBlock").innerText = code;
</script>

</body>
</html>
