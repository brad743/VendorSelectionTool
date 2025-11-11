import pandas as pd
import os

# -------------------------------
# CONFIGURATION
# -------------------------------
VENDOR_FILE = "Connected Booking System Evaluation Template (NDA).xlsx - Scoring Upload.csv"
SUMMARY_OUTPUT = "vendor_scores_summary.csv"
DETAILED_OUTPUT = "vendor_scores_detailed.csv"

# Requirement weight mapping
REQUIREMENT_WEIGHTS = {
    "critical": 4,
    "important": 3,
    "useful": 2,
    "nice to have": 1,
    "not required": 0
}

# Vendor response scoring
RESPONSE_SCORES = {
    "yes": 1,
    "no": 0,
    "not provided": 0.5
}


def normalize_case(s):
    return str(s).strip().lower()


def main():
    print("\n=== Vendor Functionality Scoring Tool ===")
    print(f"Loading fixed vendor file: {VENDOR_FILE}")

    # Load vendor functionality file
    vendor_df = pd.read_csv(VENDOR_FILE)
    vendor_df.columns = [normalize_case(c) for c in vendor_df.columns]

    # Ask for system criteria file
    criteria_path = input("\nEnter path to new system criteria CSV: ").strip()
    if not os.path.exists(criteria_path):
        print(f"Error: File not found: {criteria_path}")
        return

    criteria_df = pd.read_csv(criteria_path)
    criteria_df.columns = [normalize_case(c) for c in criteria_df.columns]

    # Prepare mappings
    function_to_req = dict(zip(criteria_df["function"], criteria_df["requirement"]))
    function_to_area = dict(zip(criteria_df["function"], criteria_df["business area"]))

    # Normalize keys
    function_to_req = {normalize_case(k): normalize_case(v) for k, v in function_to_req.items()}
    function_to_area = {normalize_case(k): v for k, v in function_to_area.items()}

    vendor_scores = []
    detailed_records = []

    # Iterate through each vendor
    for _, row in vendor_df.iterrows():
        vendor_name = row["vendor"]
        total_weight = 0
        total_score = 0
        area_scores = {}

        for func_col in vendor_df.columns:
            if func_col == "vendor":
                continue

            func_name = normalize_case(func_col)
            if func_name not in function_to_req:
                continue  # skip functions not in criteria

            req = function_to_req[func_name]
            req_weight = REQUIREMENT_WEIGHTS.get(req, 0)
            if req_weight == 0:
                continue

            response = normalize_case(row[func_col])
            response_score = RESPONSE_SCORES.get(response, 0)
            weighted_score = response_score * req_weight

            total_score += weighted_score
            total_weight += req_weight

            # Determine business area
            area = function_to_area.get(func_name, "Unspecified")
            if area not in area_scores:
                area_scores[area] = {"score": 0, "weight": 0}
            area_scores[area]["score"] += weighted_score
            area_scores[area]["weight"] += req_weight

            # Determine meets criteria
            meets = "Meets Criteria" if (response_score * 100) >= 75 else "Does Not Meet"

            detailed_records.append({
                "Vendor": vendor_name,
                "Business Area": area,
                "Function": func_col,
                "Requirement": req,
                "Response": row[func_col],
                "Weighted Score": round((weighted_score / req_weight) * 100 if req_weight > 0 else 0, 2),
                "Meets Criteria": meets
            })

        vendor_total_pct = round((total_score / total_weight) * 100 if total_weight > 0 else 0, 2)
        summary_row = {"Vendor": vendor_name, "Total Score (%)": vendor_total_pct}

        # Compute area-level averages
        for area, vals in area_scores.items():
            area_pct = round((vals["score"] / vals["weight"]) * 100 if vals["weight"] > 0 else 0, 2)
            summary_row[f"{area} (%)"] = area_pct

        vendor_scores.append(summary_row)

    # Convert to DataFrames
    summary_df = pd.DataFrame(vendor_scores)
    detailed_df = pd.DataFrame(detailed_records)

    # Save outputs
    summary_df.to_csv(SUMMARY_OUTPUT, index=False)
    detailed_df.to_csv(DETAILED_OUTPUT, index=False)

    print("\n✅ Scoring complete!")
    print(f"Summary saved to: {SUMMARY_OUTPUT}")
    print(f"Detailed results saved to: {DETAILED_OUTPUT}")
    print("\nTop Vendors by Total Score:")
    print(summary_df.sort_values("Total Score (%)", ascending=False).head(5).to_string(index=False))


if __name__ == "__main__":
    main()
