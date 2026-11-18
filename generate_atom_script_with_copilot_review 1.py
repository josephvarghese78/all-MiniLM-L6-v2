"""
ATOM Script Generator (CLI)
===========================
Reads a mapping Excel file and generates the ATOM automation script(s).

Usage:
    python generate_atom_script.py <mapping_file.xlsx> [sheet1 sheet2 ...]

If no sheets are given, sheets with Run_flag = 'Y' in Control_Sheet are used
(or all sheets except Control_Sheet if Control_Sheet is missing).

Original scripts are saved under OUTPUT_DIR_PATH/ATOM_SCRIPTS. After generation,
choose Y to review only source and target SQL with Copilot. Generated/final SQL,
JSON reviews, an Excel report, and rebuilt scripts go under OUTPUT_DIR_PATH/AI_REVIEW.
AI review requires github-copilot-sdk, sqlglot, and Copilot CLI sign-in.
"""

import asyncio
import json
import os
import re
import shutil
import subprocess
import sys
import pandas as pd
from datetime import datetime
from pathlib import Path


INPUT_FILE_PATH = r"C:\Users\rahahab\OneDrive - Manulife\ATOM_GENERATION_WITH_AI_REVIEW_1029\Mapping_Testing_1.1.xlsx"
OUTPUT_DIR_PATH = r"C:\Users\rahahab\OneDrive - Manulife\ATOM_GENERATION_WITH_AI_REVIEW_1029\output"


col_map_keys = {
    "source_table":   ["Table Name", "Source Table", "Raw Table"],
    "source_column":  ["Alias", "Source Column", "Raw Column", "Technical Name"],
    "transformation": ["Transformations", "Transformation", "SQL Logic"],
    "target_table":   ["Table name", "Target Table"],
    "target_column":  ["Field name", "Target Column", "Curated Column"]
}


# ❌ DO NOT CHANGE
def build_automation_script(table_name, s1_sql, t1_sql, output_path):
    return f"""
from datetime import date, datetime
# Initialize settings
settings_instance = Settings()
settings_instance.setsparkobject(spark)
settings_instance.projectinfo("Provide Your Project Information", "Mention EDL Layer")
settings_instance.setoutputpath("{output_path}")
settings_instance.projectname = "Provide Your Project Name"
print(len(settings_instance.projectname))
settings_instance.projectdesc = "Provide Your Project Description"
settings_instance.setwriteformat("delta")
settings_instance.setmaxerror(-1)
settings_instance.setloglevel("all")
settings_instance.replacelist = {{}}

# Initialize Writer
init_writer()

# Create TestRunner
tr = TestRunner(settings_instance)

for i in range(1):
    # Set table property
    tr.settableproperty(
        tablename="{table_name}",
        coltest=True,
        coldatatypetest=True,
        datatest=True,
        savetable=False,
        debug=True
        #excolsincoltest=["load_dt", "load_ts"],
        #excolsindatatest=["load_dt", "load_ts"]
    )
    # Initialize test
    tr.inittest()
    # Load source --- QA SCRIPT
    s1 = tr.loadsql(\"\"\"\n{s1_sql}\n    \"\"\", "source")
    # Load target
    t1 = tr.loadsql(\"\"\"\n{t1_sql}\n    \"\"\", "target")
    
    # Build tables with correct PK
    s2 = tr.buildtable("src1", "SELECT * FROM source", pk="generated_primary_key")
    t2 = tr.buildtable("trg1", "SELECT * FROM target", pk="generated_primary_key")
    # Column test
    ctr = tr.comparecolumns("source", "target")
    
    # Key test
    ktr = tr.comparekeycolumn(s2, t2)

    if ktr and ktr.get("MATCH_FOUND", 0) > 0:
        dtr = tr.comparedata(s2, t2)
        print("Data Test:", dtr)
    else:
        print("No matches found to do data testing.")

    # Finish table
    tsr = tr.finishtable()
    print("Table Summary:", tsr)

print("test completed....")

sr = tr.getsummary()
print(sr)

repname = tr.finishtest()
print(repname)

tr.printsampletestreport()
"""


def build_sql_expression(src_col, tgt_col, transform):
    src_col     = str(src_col).strip()
    tgt_col_str = "" if pd.isna(tgt_col) else str(tgt_col).strip()
    transform   = "" if pd.isna(transform) else str(transform).strip()

    if not transform or transform.upper() == "DIRECT":
        if src_col.lower() in ["", "null", "none", "nan", "na"]:
            return f"NULL AS {tgt_col_str}"
        return f"{src_col} AS {tgt_col_str}"

    parts = [p.strip().lower() for p in transform.split('+')]

    if "direct" in parts or "audit column" in parts:
        base_expr = src_col
    elif "explode_outer" in parts:
        base_expr = f"explode_outer({src_col})"
    elif "single select from value list" in parts:
        base_expr = f"array_join({src_col}, ', ')"
    else:
        base_expr = transform

    if "convert to string" in parts:
        base_expr = f"CAST({base_expr} AS STRING)"

    if tgt_col_str and re.search(rf"\bAS\s+{re.escape(tgt_col_str)}\b", str(base_expr), re.IGNORECASE):
        return base_expr

    return f"{base_expr} AS {tgt_col_str}" if tgt_col_str else base_expr


def get_sheets_to_process(file_name):
    try:
        df = pd.read_excel(file_name, sheet_name="Control_Sheet",
                           usecols=["Sheet_Name", "Run_flag"])
        df = df.dropna(subset=["Sheet_Name"])
        df["Run_flag"] = df["Run_flag"].astype(str).str.strip().str.upper()
        return df.loc[df["Run_flag"] == "Y", "Sheet_Name"].astype(str).tolist()
    except Exception:
        return [s for s in pd.ExcelFile(file_name).sheet_names if s != "Control_Sheet"]


def generate_atom_scripts(file_name, sheets_to_process, sql_reviews=None):
    scripts = []

    for sheet in sheets_to_process:
        # Detect header row
        header_row = None
        for i in range(15):
            temp_df = pd.read_excel(file_name, sheet_name=sheet, header=i)
            cols = [str(c).strip() for c in temp_df.columns]
            if sum(any(c in cols for c in v) for v in col_map_keys.values()) >= 3:
                header_row = i
                break
        if header_row is None:
            print(f"⚠️ [{sheet}] Could not detect header row — skipping.")
            continue

        pdf_mapping = pd.read_excel(file_name, sheet_name=sheet, header=header_row)

        COL_MAP = {}
        for key, names in col_map_keys.items():
            found = next((c for c in names if c in pdf_mapping.columns), None)
            if not found:
                print(f"⚠️ [{sheet}] Column for '{key}' not found — skipping.")
                break
            COL_MAP[key] = found
        else:
            select_clauses, tgt_columns = [], []
            mapping_rows = []
            source_table = None

            for _, row in pdf_mapping.iterrows():
                tgt_col = row[COL_MAP["target_column"]]
                if pd.isna(tgt_col) or str(tgt_col).strip() == "":
                    continue
                source_table = row[COL_MAP["source_table"]]
                expr = build_sql_expression(row[COL_MAP["source_column"]], tgt_col,
                                            row[COL_MAP["transformation"]])
                select_clauses.append(f"    {expr}")
                tgt_columns.append(str(tgt_col).strip())
                mapping_rows.append({
                    key: "" if pd.isna(row[column]) else str(row[column]).strip()
                    for key, column in COL_MAP.items()
                })

            if not select_clauses:
                continue

            s1_sql = (
                "SELECT\n    " +
                ",\n    ".join(select_clauses) +
                "\n    ,row_number() OVER (ORDER BY load_dt DESC) AS generated_primary_key\n"
                f"FROM {source_table}"
            )
            #t1_sql = "select\n    " + ",\n    ".join(tgt_columns) + f"\nfrom {source_table}"
            t1_sql = (
               "select\n    " +
                ",\n    ".join(tgt_columns) +
                "\n    ,row_number() OVER (ORDER BY load_dt DESC) AS generated_primary_key\n"
                f"from {source_table}"
            )

            output_path = (
                "abfss://qa-automation@cacdgfdoqaautoadls.dfs.core.windows.net/QA_Summary_Logs/"
                f"{source_table}_"
                f"{datetime.now().strftime('%d%m%Y')}_Test_Result"
            )

            scripts.append((sheet, source_table,
                            build_automation_script(source_table, s1_sql, t1_sql, output_path)))
            if sql_reviews is not None:
                sql_reviews.append({
                    "sheet": str(sheet),
                    "table": str(source_table),
                    "source_sql": s1_sql,
                    "target_sql": t1_sql,
                    "mapping_rows": mapping_rows,
                    "output_path": output_path,
                })

    return scripts


####### AI REVIEW SCRIPT START FROM HERE #######

DEFAULT_MODEL = "auto"
SQL_DIALECT = "spark"
SYSTEM_PROMPT = f"""You are a senior data engineer reviewing {SQL_DIALECT} SQL generated from a source-to-target mapping spreadsheet.
Check that target aliases, source tables/columns, transformations, NULL handling, data types, and Spark SQL syntax match the mapping. Fix only definite problems; preserve correct logic and column order. Do not invent missing requirements.
Return only a JSON object with this exact shape:
{{
  "status": "OK" or "NEEDS_FIX",
  "issues": [
    {{"severity": "HIGH|MEDIUM|LOW", "target_column": "...", "description": "...", "fix_applied": "..."}}
  ],
  "fixed_sql": "complete corrected SQL, or the original SQL if no fix is needed"
}}"""


def automated_errors(sql, mapping_rows):
    import sqlglot
    from sqlglot import exp
    from sqlglot.errors import ParseError

    try:
        tree = sqlglot.parse_one(sql, read=SQL_DIALECT)
    except ParseError as error:
        return [f"SYNTAX ERROR: {str(error)[:500]}"]

    select = tree if isinstance(tree, exp.Select) else tree.find(exp.Select)
    if select is None:
        return ["SQL does not contain a SELECT statement"]

    output_columns = {item.alias_or_name.lower() for item in select.expressions}
    errors = []
    expected_columns = [row["target_column"].lower() for row in mapping_rows]
    for column in expected_columns:
        if column not in output_columns:
            errors.append(f"MISSING TARGET COLUMN: {column}")

    duplicates = {column for column in expected_columns if expected_columns.count(column) > 1}
    for column in sorted(duplicates):
        errors.append(f"DUPLICATE TARGET COLUMN IN MAPPING: {column}")
    return errors


def parse_json_response(response_text):
    cleaned = response_text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", cleaned, flags=re.IGNORECASE)
    start = cleaned.find("{")
    end = cleaned.rfind("}")
    if start < 0 or end < start:
        raise ValueError("Copilot response did not contain a JSON object")
    result = json.loads(cleaned[start : end + 1])
    if not isinstance(result, dict) or not isinstance(result.get("fixed_sql"), str):
        raise ValueError("Copilot JSON response is missing the fixed_sql string")
    if not isinstance(result.get("issues", []), list):
        raise ValueError("Copilot JSON response has an invalid issues list")
    return result


def build_review_prompt(sheet_name, sql, mapping_rows, check_errors):
    return (
        f"Review the SQL for worksheet {sheet_name!r}.\n\n"
        f"MAPPING ROWS:\n{json.dumps(mapping_rows, indent=2)}\n\n"
        f"AUTOMATED CHECK ERRORS (must be addressed):\n"
        f"{json.dumps(check_errors, indent=2)}\n\n"
        f"SQL TO REVIEW:\n{sql}"
    )


def launch_copilot_cli():
    copilot_command = shutil.which("copilot")
    if not copilot_command:
        raise FileNotFoundError(
            "Copilot CLI was not found on PATH. Install GitHub Copilot CLI first."
        )

    if sys.platform == "win32" and Path(copilot_command).suffix.lower() in {
        ".ps1", ".cmd", ".bat",
    }:
        powershell = shutil.which("pwsh") or shutil.which("powershell")
        if not powershell:
            raise FileNotFoundError("PowerShell was not found to launch Copilot CLI.")
        return subprocess.run(
            [powershell, "-NoLogo", "-NoProfile", "-Command", "copilot"],
            check=False,
        )

    return subprocess.run([copilot_command], check=False)


async def authenticate_copilot(client):
    print("Checking Copilot login status...")
    try:
        auth = await client.get_auth_status()
    except Exception as error:
        print(f"Copilot login status could not be verified: {error}")
        return error

    while not auth.isAuthenticated:
        detail = f" ({auth.statusMessage})" if auth.statusMessage else ""
        print(f"Copilot login status: Not signed in{detail}.")
        while True:
            choice = input("Would you like to sign in to Copilot now? (Y/N): ").strip().upper()
            if choice in {"Y", "N"}:
                break
            print("Please enter Y or N.")

        if choice == "N":
            error = RuntimeError("User skipped Copilot sign-in; AI review was not run.")
            print(f"Copilot login status: Not signed in. {error}")
            return error

        print(
            "Copilot CLI is opening. At its prompt, enter /login, complete sign-in "
            "in your browser, then enter /exit to return here."
        )
        try:
            launch_copilot_cli()
        except (FileNotFoundError, OSError) as error:
            print(f"Could not launch Copilot CLI: {error}")
            print("Sign in from another terminal, then choose Y to retry or N to skip.")

        try:
            auth = await client.get_auth_status()
        except Exception as error:
            return error

    login = getattr(auth, "login", None)
    if login:
        print(f"Copilot login status: Signed in as {login}.")
    else:
        print("Copilot login status: Signed in.")
    return None


def review_report_frame(rows):
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

    return pd.DataFrame(rows).map(
        lambda value: ILLEGAL_CHARACTERS_RE.sub("", value) if isinstance(value, str) else value
    )


async def review_generated_sql(sql_reviews, review_dir):
    os.makedirs(review_dir, exist_ok=True)
    report_path = os.path.join(review_dir, "Copilot_SQL_Review_Report.xlsx")
    try:
        from copilot import CopilotClient
        from sqlglot import parse_one
    except ImportError as error:
        message = (
            f"AI review dependencies unavailable: {error}. Install with: "
            "python -m pip install github-copilot-sdk sqlglot"
        )
        review_report_frame([{"status": "NOT_REVIEWED", "error": message}]).to_excel(
            report_path, sheet_name="Summary", index=False,
        )
        print(message)
        return

    if not os.environ.get("COPILOT_CLI_PATH"):
        installed_cli = shutil.which("copilot.exe") or shutil.which("copilot")
        if installed_cli and Path(installed_cli).suffix.lower() == ".exe":
            os.environ["COPILOT_CLI_PATH"] = installed_cli

    try:
        async with CopilotClient() as client:
            auth_error = await authenticate_copilot(client)
            if auth_error is None:
                try:
                    available_models = await client.list_models()
                    print("Models available to this Copilot account:")
                    if available_models:
                        for model in available_models:
                            model_id = getattr(model, "id", "Unknown ID")
                            model_name = getattr(model, "name", None) or model_id
                            print(f"  - {model_name} ({model_id})")
                    else:
                        print("  No models were returned for this account.")
                except Exception as error:
                    print(f"Could not retrieve available Copilot models: {error}")
            else:
                print("Available models cannot be listed until Copilot sign-in succeeds.")

            print(f"Selected model setting for this review: {DEFAULT_MODEL}")
            if DEFAULT_MODEL == "auto":
                print("Copilot will automatically choose an available model for each review.")

            summaries = []
            issue_rows = []
            copilot_reviews_received = 0
            for generated in sql_reviews:
                safe_name = re.sub(r"[^\w.-]", "_", generated["sheet"])
                reviewed_path = os.path.join(review_dir, f"ATOM_{safe_name}_reviewed.py")
                final_queries = {}
                for role in ("source", "target"):
                    original_sql = generated[f"{role}_sql"]
                    expected_columns = generated["mapping_rows"] + [
                        {"target_column": "generated_primary_key"}
                    ]
                    check_errors = automated_errors(original_sql, expected_columns)
                    final_sql = original_sql
                    status = "NEEDS_HUMAN_REVIEW"
                    issues = []
                    print(f"Reviewing {generated['sheet']} {role} SQL with Copilot...")
                    try:
                        if auth_error is not None:
                            raise RuntimeError(str(auth_error))
                        session = await client.create_session(
                            model=DEFAULT_MODEL,
                            available_tools=[],
                            system_message={"mode": "append", "content": SYSTEM_PROMPT},
                        )
                        try:
                            prompt = (
                                f"Review only this {role} SQL query, not an ATOM script. "
                                "Preserve generated_primary_key and its window expression. "
                                "For source SQL, check mapping transformations and aliases. "
                                "For target SQL, select mapped target columns from the mapped "
                                "target table; do not apply source transformations.\n\n"
                                + build_review_prompt(
                                    generated["sheet"], original_sql,
                                    generated["mapping_rows"], check_errors,
                                )
                            )
                            response = await session.send_and_wait(prompt, timeout=180.0)
                            if response is None or not getattr(response.data, "content", None):
                                raise ValueError("Copilot returned an empty response")
                            review = parse_json_response(response.data.content)
                            copilot_reviews_received += 1
                        finally:
                            await session.disconnect()
                        candidate_sql = review["fixed_sql"].strip() or original_sql
                        candidate_errors = automated_errors(candidate_sql, expected_columns)
                        issues = review.get("issues", [])
                        if candidate_errors:
                            issues.append({
                                "severity": "HIGH",
                                "description": "Rejected AI SQL: " + "; ".join(candidate_errors),
                                "fix_applied": "Kept the original generated SQL",
                            })
                        else:
                            final_sql = candidate_sql
                            if final_sql != original_sql:
                                status = "FIXED"
                            elif review.get("status") == "OK" and not check_errors:
                                status = "APPROVED"
                    except Exception as error:
                        issues = [{"severity": "HIGH", "description": str(error)}]
                        print(f"AI review unavailable; keeping original {role} SQL: {error}")

                    final_queries[role] = final_sql
                    summaries.append({
                        "sheet": generated["sheet"], "role": role, "status": status,
                        "model": DEFAULT_MODEL, "issues_found": len(issues),
                        "automated_errors_before_review": "; ".join(check_errors),
                        "automated_errors_after_review": "; ".join(
                            automated_errors(final_sql, expected_columns)
                        ),
                        "original_sql": original_sql, "final_sql": final_sql,
                        "reviewed_script": reviewed_path,
                    })
                    issue_rows.extend({"sheet": generated["sheet"], "role": role, **issue}
                                      for issue in issues)

                reviewed_script = build_automation_script(
                    generated["table"], final_queries["source"], final_queries["target"],
                    generated["output_path"],
                )
                with open(reviewed_path, "w", encoding="utf-8") as script_file:
                    script_file.write(reviewed_script.strip() + "\n")

            with pd.ExcelWriter(report_path) as writer:
                review_report_frame(summaries).to_excel(writer, sheet_name="Summary", index=False)
                review_report_frame(issue_rows).to_excel(writer, sheet_name="Issues", index=False)
            total_review_requests = len(sql_reviews) * 2
            print(
                "Copilot returned review responses for "
                f"{copilot_reviews_received} of {total_review_requests} SQL queries."
            )
            print(f"AI review output saved: {review_dir}")
    except Exception as error:
        review_report_frame([{"status": "NOT_REVIEWED", "error": str(error)}]).to_excel(
            report_path, sheet_name="Summary", index=False,
        )
        print(f"AI review could not complete; original ATOM scripts are unchanged: {error}")


def main():
    file_name = sys.argv[1] if len(sys.argv) >= 2 else INPUT_FILE_PATH
    if not os.path.exists(file_name):
        print(f"❌ File not found: {file_name}")
        sys.exit(1)

    sheets = sys.argv[2:] if len(sys.argv) >= 3 else get_sheets_to_process(file_name)
    if not sheets:
        print("❌ No sheets to process.")
        sys.exit(1)

    sql_reviews = []
    scripts = generate_atom_scripts(file_name, sheets, sql_reviews=sql_reviews)
    if not scripts:
        print("❌ No ATOM scripts generated.")
        sys.exit(1)

    out_dir = os.path.join(OUTPUT_DIR_PATH, "ATOM_SCRIPTS")
    os.makedirs(out_dir, exist_ok=True)
    for sheet, table, script in scripts:
        print("\n" + "=" * 120)
        print(f"ATOM SCRIPT — Table: {table} | Sheet: {sheet}")
        print("=" * 120)
        print(script)

        safe = re.sub(r"[^\w.-]", "_", str(table))
        out_file = os.path.join(out_dir, f"ATOM_{safe}_{datetime.now():%Y%m%d_%H%M%S}.py")
        with open(out_file, "w", encoding="utf-8") as f:
            f.write(script.strip() + "\n")
        print(f"✅ Saved: {out_file}")

    while True:
        review_choice = input("Would you like AI review of source and target SQL? (Y/N): ").strip().upper()
        if review_choice in {"Y", "N"}:
            break
        print("Please enter Y or N.")
    if review_choice == "Y":
        asyncio.run(review_generated_sql(sql_reviews, os.path.join(OUTPUT_DIR_PATH, "AI_REVIEW")))
    else:
        print(f"AI review skipped. Generated ATOM script(s) are saved in: {out_dir}")


if __name__ == "__main__":
    main()