assignment_groups = combined["Assignment Group"].unique()

for ag in assignment_groups:

    if str(ag).strip() == "":
        continue

    df_group = combined[combined["Assignment Group"] == ag]

    breached_rows = df_group[df_group["__Sheet"] == "BreachedPRB"]
    about_rows = df_group[df_group["__Sheet"] == "AboutToBreachPRB"]

    if breached_rows.empty and about_rows.empty:
        continue

    # Build HTML body
    html_body = HTML_STYLE
    html_body += f"<h2>Assignment Group: {escape_html(ag)}</h2>"

    if not breached_rows.empty:
        html_body += f"<h2>EDO Open PRBs in Breach (P3–P5) – {escape_html(ag)}</h2>"
        html_body += build_html_table(breached_rows, BREACHED_COLUMNS + ["AssigneeEmail", "ManagerEmail"])

    if not about_rows.empty:
        html_body += f"<h2>EDO Open PRBs about to Breach (P3–P5) – {escape_html(ag)}</h2>"
        html_body += build_html_table(about_rows, ABOUT_COLUMNS + ["AssigneeEmail", "ManagerEmail"])

    # --- APPLY ITSM TO/CC LOGIC ---
    TO_LIST, CC_LIST_FINAL = determine_to_cc(df_group)

    # If no valid TO found, default to your email (safe)
    if not TO_LIST:
        TO_LIST = ["shivramtripathi700@gmail.com"]

    print(f"\nSending email for AG: {ag}")
    print("TO :", TO_LIST)
    print("CC :", CC_LIST_FINAL)

    send_email(ag, html_body, TO_LIST, CC_LIST_FINAL)
