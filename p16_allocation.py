#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
P16 floorball allocation script
- Strict max 1 GK per player (default; configurable)
- Base caps: max 2 home + max 2 away standard call-ups (configurable)
- Third line (RESERVE) only created if exactly 4 players available for it (configurable)
- Adds an info-only row "MÖJLIGA RESERVER" with all other available players (does not affect counts)
- Overview columns on the main sheet: Kallelser Hemma/Borta/Totalt, Reservkallelser, Målvaktsgånger

Usage:
    python p16_allocation.py INPUT.xlsx OUTPUT.xlsx [--sheet "Formulärsvar 1 (exakt)"]

Optional flags:
    --max-home-base 2
    --max-away-base 2
    --gk-cap 1
    --require-exact-reserve-four (if set, create KEDJA 3 only with exactly 4 players; default True)
    --no-require-exact-reserve-four (create KEDJA 3 when at least 4 players; default False)
    --prefer-gk-volunteers (default True)
"""
import argparse, sys, re
from collections import defaultdict, OrderedDict
from typing import List, Dict, Any

import pandas as pd

def yn(val) -> bool:
    s = str(val).strip().lower()
    return s in {"ja", "j", "yes", "y", "1", "true"}

def find_match_columns(df: pd.DataFrame) -> List[str]:
    return [c for c in df.columns if "(" in c and ")" in c and ("Hemma" in c or "Borta" in c)]

def infer_vill_borta(df: pd.DataFrame, match_cols: List[str]) -> pd.Series:
    # Prefer '#Borta svar' if present; otherwise, count "Ja" in away match columns
    if "#Borta svar" in df.columns:
        return (df["#Borta svar"].fillna(0) > 0)
    away_cols = [c for c in match_cols if "Borta" in c]
    if not away_cols:
        return pd.Series([False]*len(df), index=df.index)
    return df[away_cols].applymap(yn).sum(axis=1) > 0

def sanitize_sheet_name(name: str) -> str:
    safe = re.sub(r'[:\\/?*\[\]]', '-', name)
    if len(safe) > 31:
        safe = safe[:28] + "..."
    return safe

def allocate(
    df_in: pd.DataFrame,
    sheet_name: str,
    max_home_base: int = 2,
    max_away_base: int = 2,
    gk_cap: int = 1,
    require_exact_reserve_four: bool = True,
    prefer_gk_volunteers: bool = True
) -> Dict[str, Any]:
    df = df_in.copy()

    # Normalize inputs
    for col in ["Målvakt", "Reserv"]:
        if col not in df.columns:
            raise ValueError(f"Kunde inte hitta kolumnen '{col}' i bladet '{sheet_name}'.")
    df["Målvakt_bool"] = df["Målvakt"].map(yn)
    df["Reserv_bool"] = df["Reserv"].map(yn)

    match_cols = find_match_columns(df)
    if not match_cols:
        raise ValueError("Hittade inga matchkolumner. Kontrollera att rubrikerna innehåller '(Hemma)' eller '(Borta)'.")

    df["Vill_borta"] = infer_vill_borta(df, match_cols)

    player_rows = list(df.index)
    player_names = df["Barnets namn"].tolist() if "Barnets namn" in df.columns else [str(i) for i in player_rows]

    # Counters
    base_total = defaultdict(int)      # standardkallelser totalt
    base_home = defaultdict(int)       # standardkallelser hemma
    base_away = defaultdict(int)       # standardkallelser borta
    home_total = defaultdict(int)      # alla hemmamatcher spelade (inkl. reserv)
    away_total = defaultdict(int)      # alla bortamatcher spelade (inkl. reserv)
    reserve_calls = defaultdict(int)   # reservkallelser
    gk_assign = defaultdict(int)       # målvaktsgånger
    gk_cap_violations = []

    def has_base_capacity(i: int, match_typ: str) -> bool:
        if match_typ == "Hemma":
            return (base_home[i] < max_home_base) and (base_total[i] < (max_home_base + max_away_base))
        else:
            return (base_away[i] < max_away_base) and (base_total[i] < (max_home_base + max_away_base))

    def pref_key_for_pool(i: int) -> tuple:
        # Prefer volunteers if configured; then fewer standardkallelser; then fewer reservkallelser;
        # then away-willing; then stable player id (column 'Spelare' if present, else index)
        prefer = 0 if (df.loc[i, "Målvakt_bool"] if prefer_gk_volunteers else False) else 1
        stable_id = int(df.loc[i, "Spelare"]) if "Spelare" in df.columns else int(i)+1
        return (prefer, base_total[i], reserve_calls[i], -int(df.loc[i,"Vill_borta"]), stable_id)

    def fairness_key_field(i: int, match_typ: str) -> tuple:
        stable_id = int(df.loc[i, "Spelare"]) if "Spelare" in df.columns else int(i)+1
        return (base_total[i], -int(df.loc[i,"Vill_borta"]), base_total[i] + reserve_calls[i], stable_id)

    per_match = OrderedDict()

    for col in match_cols:
        match_typ = "Hemma" if "Hemma" in col else "Borta"
        avail = [i for i in player_rows if yn(df.loc[i, col])]

        # GK selection with strict cap if possible.
        pools = [
            [i for i in avail if gk_assign[i] < gk_cap and has_base_capacity(i, match_typ)],
            [i for i in avail if gk_assign[i] < gk_cap],
        ]
        chosen_gk: List[int] = []
        for pool in pools:
            if len(chosen_gk) < 2:
                for i in sorted(pool, key=pref_key_for_pool):
                    if i not in chosen_gk:
                        chosen_gk.append(i)
                    if len(chosen_gk) == 2:
                        break

        if len(chosen_gk) < 2:
            # Last resort: exceed GK cap (should be rare)
            fallback = [i for i in avail if i not in chosen_gk]
            for i in sorted(fallback, key=pref_key_for_pool):
                if i not in chosen_gk:
                    chosen_gk.append(i)
                    if gk_assign[i] >= gk_cap:
                        gk_cap_violations.append((col, player_names[i]))
                if len(chosen_gk) == 2:
                    break

        gks = chosen_gk[:2]

        # Field base two lines (8 players) with base capacity
        field_base_pool = [i for i in avail if i not in gks and has_base_capacity(i, match_typ)]
        field_base_sorted = sorted(field_base_pool, key=lambda i: fairness_key_field(i, match_typ))
        field_base = field_base_sorted[:8]

        # If fewer than 8, pull from reserves (volunteers)
        if len(field_base) < 8:
            extra_reserve_pool = [i for i in avail if i not in gks and i not in field_base and df.loc[i, "Reserv_bool"]]
            extra_reserve_sorted = sorted(extra_reserve_pool, key=lambda i: (reserve_calls[i], base_total[i] + reserve_calls[i], -int(df.loc[i,"Vill_borta"]), int(df.loc[i,"Spelare"]) if "Spelare" in df.columns else i+1))
            need = 8 - len(field_base)
            field_base += extra_reserve_sorted[:need]

        # Optional third reserve line
        remaining_pool = [i for i in avail if i not in gks and i not in field_base and df.loc[i, "Reserv_bool"]]
        reserve_sorted = sorted(remaining_pool, key=lambda i: (reserve_calls[i], base_total[i] + reserve_calls[i], -int(df.loc[i,"Vill_borta"]), int(df.loc[i,"Spelare"]) if "Spelare" in df.columns else i+1))
        if require_exact_reserve_four:
            reserve_line = reserve_sorted[:4] if len(reserve_sorted) >= 4 else []
        else:
            reserve_line = reserve_sorted[:4]

        # Possible reserves (info-only)
        assigned_idxs = set(gks + field_base + (reserve_sorted[:4] if (not require_exact_reserve_four) or (len(reserve_sorted) >= 4) else []))
        possible_reserves = [player_names[i] for i in sorted([i for i in avail if i not in assigned_idxs], key=lambda i: int(df.loc[i,"Spelare"]) if "Spelare" in df.columns else i+1)]

        per_match[col] = {
            "typ": match_typ,
            "gk": [player_names[i] for i in gks],
            "line1": [player_names[i] for i in field_base[:4]],
            "line2": [player_names[i] for i in field_base[4:8]],
            "reserve_line": [player_names[i] for i in reserve_line],
            "possible_reserves": possible_reserves
        }

        # Update counters for assigned players ONLY (possible_reserves do not affect counts)
        for i in gks:
            gk_assign[i] += 1
            is_base_now = has_base_capacity(i, match_typ)
            if match_typ == "Hemma":
                home_total[i] += 1
                if is_base_now:
                    base_home[i] += 1
                    base_total[i] += 1
                else:
                    reserve_calls[i] += 1
            else:
                away_total[i] += 1
                if is_base_now:
                    base_away[i] += 1
                    base_total[i] += 1
                else:
                    reserve_calls[i] += 1

        for i in field_base:
            is_base_now = has_base_capacity(i, match_typ)
            if match_typ == "Hemma":
                home_total[i] += 1
                if is_base_now:
                    base_home[i] += 1
                    base_total[i] += 1
                else:
                    reserve_calls[i] += 1
            else:
                away_total[i] += 1
                if is_base_now:
                    base_away[i] += 1
                    base_total[i] += 1
                else:
                    reserve_calls[i] += 1

        for i in reserve_line:
            if match_typ == "Hemma":
                home_total[i] += 1
            else:
                away_total[i] += 1
            reserve_calls[i] += 1

    # Build output DF with overview columns
    out = df.copy()
    out["Kallelser Hemma"] = [base_home[i] for i in player_rows]
    out["Kallelser Borta"] = [base_away[i] for i in player_rows]
    out["Kallelser Totalt"] = [base_home[i] + base_away[i] for i in player_rows]
    out["Reservkallelser"] = [reserve_calls[i] for i in player_rows]
    out["Målvaktsgånger"] = [gk_assign[i] for i in player_rows]
    out["Antal Hemma matcher"] = [home_total[i] for i in player_rows]
    out["Antal Borta matcher"] = [away_total[i] for i in player_rows]

    # Column order for the main sheet
    match_cols = find_match_columns(df_in)
    base_cols = ["Spelare","Barnets namn","Målvakt","Reserv"]
    overview_cols = ["Kallelser Hemma","Kallelser Borta","Kallelser Totalt","Reservkallelser","Målvaktsgånger"]
    reference_cols = ["Antal Hemma matcher","Antal Borta matcher"]
    final_cols = [c for c in base_cols if c in out.columns] + overview_cols + reference_cols + match_cols
    out = out[final_cols]

    return {
        "main": out,
        "per_match": per_match,
        "gk_cap_violations": gk_cap_violations
    }

def write_excel(result: Dict[str, Any], output_path: str):
    main_df: pd.DataFrame = result["main"]
    per_match = result["per_match"]

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        main_df.to_excel(writer, sheet_name="Formulärsvar 1 (exakt)", index=False)
        used_names = set()
        for mcol, info in per_match.items():
            width = 1 + max(
                len(info["gk"]),
                len(info["line1"]),
                len(info["line2"]),
                len(info["reserve_line"]),
                len(info["possible_reserves"]),
            )
            def pad(row):
                return row + [""] * (width - (len(row)-1))

            safe_name = sanitize_sheet_name(mcol)
            base_name = safe_name
            n = 1
            while safe_name in used_names:
                n += 1
                safe_name = (base_name[:27] + f"-{n}")
            used_names.add(safe_name)

            rows = []
            rows.append(pad(["MÅLVAKTER (max 1)"] + info["gk"]))
            rows.append(pad(["KEDJA 1 (UTE)"] + info["line1"]))
            rows.append(pad(["KEDJA 2 (UTE)"] + info["line2"]))
            if len(info["reserve_line"]) == 4:
                rows.append(pad(["KEDJA 3 (RESERV)"] + info["reserve_line"]))
            rows.append(pad(["MÖJLIGA RESERVER"] + info["possible_reserves"]))

            pd.DataFrame(rows).to_excel(writer, sheet_name=safe_name, index=False, header=False)

def main():
    parser = argparse.ArgumentParser(description="P16 allocation script")
    parser.add_argument("input", help="Input Excel (med fliken 'Formulärsvar 1 (exakt)')")
    parser.add_argument("output", help="Output Excel")
    parser.add_argument("--sheet", default="Formulärsvar 1 (exakt)", help="Bladnamn (default: 'Formulärsvar 1 (exakt)')")
    parser.add_argument("--max-home-base", type=int, default=2, help="Max standardkallelser hemma per spelare")
    parser.add_argument("--max-away-base", type=int, default=2, help="Max standardkallelser borta per spelare")
    parser.add_argument("--gk-cap", type=int, default=1, help="Max målvaktsmatcher per spelare")
    parser.add_argument("--require-exact-reserve-four", dest="require_exact_reserve_four", action="store_true", default=True, help="Skapa KEDJA 3 (RESERV) endast om exakt 4")
    parser.add_argument("--no-require-exact-reserve-four", dest="require_exact_reserve_four", action="store_false", help="Skapa KEDJA 3 när minst 4 finns")
    parser.add_argument("--prefer-gk-volunteers", dest="prefer_gk_volunteers", action="store_true", default=True, help="Prioritera spelare som vill vara målvakt")
    parser.add_argument("--no-prefer-gk-volunteers", dest="prefer_gk_volunteers", action="store_false", help="Ignorera GK-preferens vid val")
    args = parser.parse_args()

    # Read input
    try:
        df = pd.read_excel(args.input, sheet_name=args.sheet)
    except Exception as e:
        print(f"Kunde inte läsa '{args.input}' / blad '{args.sheet}': {e}", file=sys.stderr)
        sys.exit(1)

    try:
        result = allocate(
            df_in=df,
            sheet_name=args.sheet,
            max_home_base=args.max_home_base,
            max_away_base=args.max_away_base,
            gk_cap=args.gk_cap,
            require_exact_reserve_four=args.require_exact_reserve_four,
            prefer_gk_volunteers=args.prefer_gk_volunteers
        )
        write_excel(result, args.output)
    except Exception as e:
        print(f"Fel vid fördelning: {e}", file=sys.stderr)
        sys.exit(2)

    if result["gk_cap_violations"]:
        print("VARNING: Några matcher krävde att GK-taket överträddes:", file=sys.stderr)
        for col, name in result["gk_cap_violations"]:
            print(f"  {col}: {name}", file=sys.stderr)
    else:
        print("Klart! Inga överträdelser av GK-taket. Resultatet finns i:", args.output)

if __name__ == "__main__":
    main()
