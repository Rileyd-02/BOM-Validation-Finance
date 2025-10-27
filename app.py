import streamlit as st
import pandas as pd
from thefuzz import fuzz
from io import BytesIO

st.set_page_config(page_title="SAP vs PLM Validation Tool", layout="wide")
st.title("📊 SAP vs PLM Validation Tool")

st.write("""
Upload your SAP (Base) file and PLM file.  
The tool checks:
- Material matches  
- Component issues (`'-'` or starting with '3')  
- Vendor Reference matches (exact or inside description)  
- Adjusted consumption comparison (Base Qty handled as decimal)  
and provides summary counts, duplicates report, and a clean preview.
""")

# ------------------------
# File Uploads
# ------------------------
sap_file = st.file_uploader("📤 Upload SAP Excel File", type=["xlsx"])
plm_file = st.file_uploader("📤 Upload PLM Excel File", type=["xlsx"])

if sap_file and plm_file:
    try:
        sap_df = pd.read_excel(sap_file)
        plm_df = pd.read_excel(plm_file)

        # Clean column names
        sap_df.columns = sap_df.columns.str.strip()
        plm_df.columns = plm_df.columns.str.strip()

        # --- Duplicates Report ---
        sap_dupes = sap_df[sap_df.duplicated(subset=["Material"], keep=False)]
        plm_dupes = plm_df[plm_df.duplicated(subset=["Material"], keep=False)]

        with st.expander("🔁 Duplicates Report"):
            st.write("**SAP duplicates (by Material):**", len(sap_dupes))
            if not sap_dupes.empty:
                st.dataframe(sap_dupes)
            st.write("**PLM duplicates (by Material):**", len(plm_dupes))
            if not plm_dupes.empty:
                st.dataframe(plm_dupes)

        # --- Step 1: Material Merge ---
        merged_df = pd.merge(
            sap_df, plm_df,
            on="Material",
            how="left",
            suffixes=("_SAP", "_PLM")
        )

        merged_df["Material_Match"] = merged_df["Material"].notna().map(
            {True: "Matched", False: "Missing in PLM"}
        )

        # --- Step 2: Component Flag ---
        merged_df["Component_Flag"] = merged_df["Component"].apply(
            lambda x: "Check (Invalid)" if isinstance(x, str) and (x.startswith("3") or "-" in x) else "OK"
        )

        # --- Step 3: Vendor Reference Logic ---
        def check_vendor_ref(row):
            plm_ref = str(row.get("Vendor Reference_PLM", "")).strip()
            sap_v_ref = str(row.get("Vendor Reference_SAP", "")).strip()
            sap_desc = str(row.get("Material Description", "")).strip()

            if not plm_ref:
                return "No Vendor Ref in PLM"
            if plm_ref == sap_v_ref:
                return "Exact Match"
            if plm_ref in sap_desc:
                return "Found in Description"
            return "Not Found"

        merged_df["VendorRef_Status"] = merged_df.apply(check_vendor_ref, axis=1)

        # --- Step 4: Adjusted SAP Consumption (Base Quantity handled) ---
        def calculate_sap_consumption(row):
            try:
                comp_qty = float(row.get("Comp.Qty.", 0))
                base_qty = float(row.get("Base Quantity(AA)", 1))
                if base_qty in [1000, 100, 1] and base_qty != 0:
                    return round(comp_qty / base_qty, 4)
                return 0.0
            except Exception:
                return 0.0

        merged_df["SAP_Consumption"] = merged_df.apply(calculate_sap_consumption, axis=1)

        # --- Step 5: PLM Consumption ---
        if "Qty(Cons.)" in plm_df.columns:
            merged_df["PLM_Consumption"] = merged_df["Qty(Cons.)"].fillna(0).astype(float).round(4)
        else:
            merged_df["PLM_Consumption"] = 0.0

        # --- Step 6: Difference and Status ---
        merged_df["Consumption_Difference"] = (merged_df["SAP_Consumption"] - merged_df["PLM_Consumption"]).round(4)
        merged_df["Consumption_Status"] = merged_df.apply(
            lambda x: "SAP Consumption Higher" if x["Consumption_Difference"] > 0 else "OK",
            axis=1
        )

        # --- Step 7: Similarity Scores ---
        def safe_ratio(a, b):
            return fuzz.token_sort_ratio(str(a), str(b))

        merged_df["Material_Similarity"] = merged_df.apply(
            lambda x: safe_ratio(x.get("Material", ""), x.get("Material", "")), axis=1
        )
        merged_df["Color_Similarity"] = merged_df.apply(
            lambda x: safe_ratio(x.get("Color_SAP", ""), x.get("Color_PLM", "")), axis=1
        )

        # --- Step 8: Summary (deduplicated) ---
        summary_df = merged_df.drop_duplicates(subset=["Material"])
        total_rows = len(summary_df)
        matched_materials = (summary_df["Material_Match"] == "Matched").sum()
        invalid_components = (summary_df["Component_Flag"] == "Check (Invalid)").sum()
        sap_higher = (summary_df["Consumption_Status"] == "SAP Consumption Higher").sum()

        st.subheader("📋 Summary Overview")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Records", total_rows)
        c2.metric("Material Matches", matched_materials)
        c3.metric("Invalid Components", invalid_components)
        c4.metric("SAP Higher Consumption", sap_higher)

        # --- Step 9: Format decimals for output ---
        decimal_cols = ["SAP_Consumption", "PLM_Consumption", "Consumption_Difference"]
        for col in decimal_cols:
            if col in merged_df.columns:
                merged_df[col] = merged_df[col].astype(float).map("{:.4f}".format)

        # --- Step 10: Save to Excel ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            merged_df.to_excel(writer, sheet_name="Comparison_Report", index=False)
            if not sap_dupes.empty:
                sap_dupes.to_excel(writer, sheet_name="SAP_Duplicates", index=False)
            if not plm_dupes.empty:
                plm_dupes.to_excel(writer, sheet_name="PLM_Duplicates", index=False)
        output.seek(0)

        st.download_button(
            label="📥 Download Full Comparison Report",
            data=output,
            file_name="SAP_PLM_Comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # --- Step 11: Preview ---
        st.subheader("🔍 Preview of Results")
        preview_cols = [
            "Bill of material", "Material", "Material Description",
            "Vendor Reference_SAP", "Vendor Reference_PLM",
            "Component", "SAP_Consumption", "PLM_Consumption",
            "Consumption_Difference", "Material_Match", "Component_Flag",
            "VendorRef_Status", "Consumption_Status",
            "Material_Similarity", "Color_Similarity"
        ]
        preview_cols = [c for c in preview_cols if c in merged_df.columns]
        st.dataframe(summary_df[preview_cols].head(100))

    except Exception as e:
        st.error(f"❌ Error while processing: {e}")
else:
    st.info("⬆️ Please upload both SAP and PLM files to start comparison.")
