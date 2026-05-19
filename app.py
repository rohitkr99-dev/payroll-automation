import streamlit as st
import pandas as pd
from datetime import datetime, time
from io import BytesIO

st.title("Payroll Automation System")

uploaded_file = st.file_uploader(
    "Upload attendance file",
    type=["xls", "xlsx"]
)

def process_payroll(file):

    raw = pd.read_excel(file, header=None)

    employees = []
    current_emp = None

    for _, row in raw.iterrows():

        first = row[0]

        # employee row
        if isinstance(first, str) and "Mr." in first:

            parts = first.strip().split(" ", 1)

            emp_id = parts[0]
            emp_name = parts[1] if len(parts) > 1 else ""

            current_emp = {
                "Employee ID": emp_id,
                "Employee Name": emp_name,
                "days": {}
            }

            employees.append(current_emp)
            continue

        # date row
        if current_emp and isinstance(first, str) and "/" in first:

            try:
                dt = datetime.strptime(
                    first.strip(),
                    "%d/%b/%Y"
                ).date()

            except:
                continue

            punches = []

            for col in [1, 2, 3, 4]:

                val = row[col]

                if pd.notna(val):

                    try:
                        punches.append(
                            pd.to_datetime(val).time()
                        )

                    except:
                        pass

            current_emp["days"][dt.day] = punches

    summary_rows = []
    issue_rows = []

    for emp in employees:

        for row_type in ["NH", "FOT", "EOT", "DETAILS"]:

            row = {
                "Employee ID": emp["Employee ID"],
                "Employee Name": emp["Employee Name"],
                "Type": row_type
            }

            for d in range(1, 16):

                punches = emp["days"].get(d, [])

                value = ""

                if len(punches) > 0:

                    first_in = punches[0]
                    last_out = punches[-1]

                    issues = []

                    # single punch handling
                    if len(punches) == 1:

                        if first_in < time(12, 0):

                            issues.append(
                                f"Single Morning Punch {first_in.strftime('%H:%M')}"
                            )

                            last_out = time(18, 0)

                        else:

                            issues.append(
                                f"Single Evening Punch {first_in.strftime('%H:%M')}"
                            )

                            first_in = time(8, 0)

                    # early in
                    if first_in <= time(7, 0):

                        issues.append(
                            f"Early IN {first_in.strftime('%H:%M')}"
                        )

                        first_in = time(8, 0)

                    # extra ot
                    eot = 0

                    out_minutes = (
                        last_out.hour * 60 +
                        last_out.minute
                    )

                    if out_minutes >= 19 * 60:

                        eot = int(
                            (out_minutes - 18 * 60) // 60
                        )

                    # row types
                    if row_type == "NH":
                        value = 8

                    elif row_type == "FOT":
                        value = 1

                    elif row_type == "EOT":
                        value = eot

                    elif row_type == "DETAILS":

                        value = (
                            f"{first_in.strftime('%H:%M')} - "
                            f"{last_out.strftime('%H:%M')}"
                        )

                        if issues:
                            value += " | " + "; ".join(issues)

                            issue_rows.append({
                                "Employee ID": emp["Employee ID"],
                                "Employee Name": emp["Employee Name"],
                                "Date": d,
                                "Issue": "; ".join(issues)
                            })

                row[str(d)] = value

            summary_rows.append(row)

    summary_df = pd.DataFrame(summary_rows)

    issues_df = pd.DataFrame(issue_rows)

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        summary_df.to_excel(
            writer,
            sheet_name="Payroll Summary",
            index=False
        )

        issues_df.to_excel(
            writer,
            sheet_name="Collective Issues",
            index=False
        )

    output.seek(0)

    return output

if uploaded_file:

    st.success("File uploaded successfully")

    result = process_payroll(uploaded_file)

    st.download_button(
        label="Download Payroll Report",
        data=result,
        file_name="payroll_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
