@st.cache_data(ttl=REFRESH)
def parse(fb):
    raw = pd.read_excel(BytesIO(fb), sheet_name=SHEET_NAME, header=None, engine="openpyxl")

    # Use row 1 as the actual header row, exactly like your Excel/dashboard structure
    header_row = 1

    def norm(x):
        return "".join(ch for ch in str(x).lower().strip() if ch.isalnum())

    headers = {}
    for c in range(raw.shape[1]):
        h = raw.iloc[header_row, c]
        if not pd.isna(h):
            headers[norm(h)] = c

    def col(possible_names, fallback):
        for name in possible_names:
            key = norm(name)
            if key in headers:
                return headers[key]
        return fallback

    C_SNO = col(["S_No"], 0)
    C_JOB = col(["Job_Ref"], 2)
    C_LPO = col(["LPO_Ref"], 3)
    C_CUSTOMER = col(["Customer"], 4)
    C_PROJECT = col(["Project_Name"], 5)
    C_REGION = col(["Region"], 6)
    C_LOCATION = col(["Location"], 7)
    C_WORK = col(["Work _Status", "Work_Status", "Work Status"], 25)
    C_MATERIAL = col(["MATERIAL STATUS", "Material_Status", "Material Status"], 27)
    C_PROGRESS = col(["Overall Progress %", "PROGRESS", "Progress"], 28)
    C_PRIORITY = col(["Priority"], 29)
    C_STATUS = col(["Status"], 30)
    C_ENG = col(["Engineering %", "Engineering_Pct"], 31)
    C_DEL = col(["Delivery%", "Delivery %", "Delivery_Pct"], 32)
    C_EXEC = col(["Execution %", "Execution_Pct"], 33)

    def safe_num(x, default=0.0):
        try:
            if pd.isna(x):
                return default
            if isinstance(x, str):
                x = x.strip().replace("%", "")
                if not x:
                    return default
            return round(float(x), 1)
        except Exception:
            return default

    rows = []

    for i in range(2, len(raw)):
        r = raw.iloc[i]

        if (pd.isna(r[C_SNO]) or str(r[C_SNO]).strip() == "") and (
            pd.isna(r[C_PROJECT]) or str(r[C_PROJECT]).strip() == ""
        ):
            continue

        try:
            sno = str(int(float(r[C_SNO])))
        except Exception:
            sno = _v(r[C_SNO])

        rows.append({
            "s_no": sno,
            "job_ref": _v(r[C_JOB], "N/A"),
            "lpo_ref": _v(r[C_LPO], "N/A"),
            "customer": _v(r[C_CUSTOMER], "Unknown"),
            "project_name": _v(r[C_PROJECT]),
            "region": _v(r[C_REGION], "Unknown"),
            "location": _v(r[C_LOCATION]),
            "work_status": _v(r[C_WORK]),
            "material_status": _v(r[C_MATERIAL], "Not Ordered"),
            "progress": safe_num(r[C_PROGRESS]),
            "priority": _v(r[C_PRIORITY], "Medium") or "Medium",
            "status": _v(r[C_STATUS], "Not Started"),
            "eng_pct": safe_num(r[C_ENG]),
            "del_pct": safe_num(r[C_DEL]),
            "exec_pct": safe_num(r[C_EXEC]),
            "del_vals": [_v(r[c]) for c in range(14, 25)],
        })

    return rows
