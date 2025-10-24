from __future__ import annotations
import pandas as pd
import numpy as np
import yaml

TYPE_MAP = {
    "string": "string",
    "int": "int64",
    "float": "float64",
    "date": "datetime64[ns]",
}

def _parse_type(spec: str):
    # "int>=0" → ("int", (">=", 0.0)) | "date" → ("date", None)
    if ">=" in spec:
        base, val = spec.split(">=", 1)
        return base.strip(), (">=", float(val))
    if "<=" in spec:
        base, val = spec.split("<=", 1)
        return base.strip(), ("<=", float(val))
    return spec.strip(), None

def _coerce_and_check_types(df: pd.DataFrame, column_types: dict):
    problems = []
    for col, spec in column_types.items():
        if col not in df.columns:
            problems.append({"name": "missing_column", "level": "error", "detail": col})
            continue

        base, cond = _parse_type(spec)

        # casting
        if base == "date":
            df[col] = pd.to_datetime(df[col], errors="coerce")
            bad = df[col].isna()
            if bad.any():
                problems.append({"name": f"{col}_type", "level": "error",
                                 "detail": f"invalid dates in {bad.sum()} rows"})
        else:
            if base in ("int", "float"):
                df[col] = pd.to_numeric(df[col], errors="coerce")
            try:
                df[col] = df[col].astype(TYPE_MAP.get(base, "object"))
            except Exception:
                problems.append({"name": f"{col}_type", "level": "error",
                                 "detail": f"cannot cast to {base}"})

        # simple range condition if defined
        if cond and col in df.columns:
            op, val = cond
            if op == ">=":
                bad = df[col] < val
            elif op == "<=":
                bad = df[col] > val
            else:
                bad = pd.Series(False, index=df.index)
            if bad.any():
                problems.append({"name": f"{col}_range", "level": "error",
                                 "detail": f"{int(bad.sum())} values violate {base}{op}{val}"})
    return df, problems

def _days_since(series: pd.Series) -> pd.Series:
    # assumes datetime64 already
    return (pd.Timestamp.now(normalize=True) - series.dt.normalize()).dt.days

def _eval_expr(df: pd.DataFrame, expr: str) -> pd.Series:
    # sandboxed eval for simple arithmetic/column ops
    env = {"np": np, "days_since": _days_since}
    return df.eval(expr, engine="python", local_dict=env)

def _run_checks(df: pd.DataFrame, checks: list[dict]) -> list[dict]:
    findings = []
    for chk in checks or []:
        try:
            mask = ~_eval_expr(df, chk["expr"])
            if mask.any():
                findings.append({
                    "name": chk.get("name", chk["expr"]),
                    "level": chk.get("level", "error"),
                    "failed_rows": int(mask.sum())
                })
        except Exception as e:
            findings.append({
                "name": chk.get("name", "invalid_check"),
                "level": "error",
                "detail": f"failed to evaluate expr: {chk.get('expr')} ({e})"
            })
    return findings

def _validate_sheet(df: pd.DataFrame, spec: dict):
    problems = []

    # required columns
    required = spec.get("required_columns", [])
    missing = [c for c in required if c not in df.columns]
    if missing:
        problems.append({"name":"missing_columns","level":"error","detail":",".join(missing)})
        return problems  # stop early if required columns are missing

    # types + ranges
    _, type_problems = _coerce_and_check_types(df, spec.get("column_types", {}))
    problems.extend(type_problems)

    # row-level checks
    problems.extend(_run_checks(df, spec.get("checks", [])))
    return problems

def validate_workbook(xl: pd.ExcelFile, rules_path: str):
    """
    xl: pd.ExcelFile (already opened in the app)
    rules_path: path to rules.yaml
    """
    with open(rules_path, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}

    results = []
    for sheet_spec in cfg.get("sheets", []):
        name = sheet_spec["name"]
        if name not in xl.sheet_names:
            results.append({
                "sheet": name,
                "ok": False,
                "problems": [{"name":"missing_sheet","level":"error","detail":"sheet not found"}]
            })
            continue

        df = xl.parse(name)
        problems = _validate_sheet(df, sheet_spec)
        ok = not any(p.get("level") == "error" for p in problems)
        results.append({"sheet": name, "ok": ok, "problems": problems})

    return {"version": cfg.get("version", 1), "results": results}
