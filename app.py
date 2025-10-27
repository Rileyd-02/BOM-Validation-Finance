import streamlit as st
import pandas as pd
from thefuzz import fuzz
from io import BytesIO

# ------------------------
# Page Setup
# ------------------------
st.set_page_config(page_title="SAP vs PLM Comparison", layout="wide")
st.title("📊 SAP vs PLM Validation Tool")

st.write("""
Upload your **SAP (Base)** file and **PLM** file.  
This tool:
- Detects duplicate Materials  
- Validates Components (`'-'` or starting with '3'`)  
- Matches Vendor References (direct or in description)  
- Adjusts SAP Consumption by Base Quantity (1000, 100, or 1)  
- Compares normalized SAP vs PLM consumption  
- Calculates similarity scores  
- Provides accurate summary counts without duplication inflation  
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

        # ------------------------
        # Duplicate Report
        # ------------------------
        sap_dupes = sap_df[sap_df.duplicated(subset=["Material"], keep=False)]
        plm_dupes = plm_df[plm_df.duplicated(subset=["Material"], keep=False)]

        with st.expander("📋 View Duplicates Report"):
            c1, c2 = st.columns(2)
            c1.subheader("SAP Duplicates")
            if len(sap_dupes) > 0:
                c1.dataframe(sap_dupes)
            else:
                c1.success("✅ No duplicate Materials in SAP file.")

            c2.subheader("PLM Duplicates")
            if len(plm_dupes) > 0:
                c2.dataframe(plm_dupes)
            else:
                c2.success("✅ No duplicate Materials in PLM file.")

        # ------------------------
        # Step 1: Merge
        # ------------------------
        merged_df = pd.merge(
            sap_df, plm_df,
            on="Material",
            how="left",
            suffixes=("_SAP", "_PLM")
        )
        merged_df["Material_Match"] = merged_df["Material"].notna().map({True: "Matched", False: "Missing in PLM"})

        # ------------------------
        # Step 2: Component Flag
        # ------------------------
        if "Component" in merged_df.columns:
            merged_df["Component_Flag"] = merged_df["Component"].apply(
                lambda x: "Check (Invalid)" if isinstance(x, str) and (x.startswith("3") or "-" in x) else "OK"
            )
        else:
            merged_df["Component_Flag"] = "Column Missing"

        # ------------------------
        # Step 3: Vendor Reference Match
        # ------------------------
        def check_vendor_ref(row):
            plm_ref = str(row.get("Vendor Reference_PLM", "")).strip()
            sap_v_ref = str(row.get("Vendor Reference_SAP", "")).strip()
            sap_desc = str(row.get("Material Description_SAP", "")).strip()

            if not plm_ref:
                return "No Vendor Ref in PLM"
            if plm_ref == sap_v_ref:
                return "Exact Match"
            if plm_ref in sap_desc:
                return "Found in Description"
            return "Not Found"

        merged_df["VendorRef_Status"] = merged_df.apply(check_vendor_ref, axis=1)

        # ------------------------
        # Step 4: Normalize SAP Consumption by Base Quantity
        # ------------------------
        sap_qty_col = next((c for c in sap_df.columns if "Comp" in c and "Qty" in c), None)
        base_qty_col = next((c for c in sap_df.columns if "Base" in c and "Qty" in c), None)
        plm_qty_col = next((c for c in plm_df.columns if "Qty" in c and "Cons" in c), None)

        merged_df["SAP_Consumption"] = 0.0
        merged_df["PLM_Consumption"] = 0.0

        if sap_qty_col:
            merged_df["SAP_Consumption"] = merged_df[sap_qty_col].fillna(0)
            if base_qty_col:
                merged_df["Base_Qty"] = merged_df[base_qty_col].fillna(1)
                merged_df["SAP_Consumption"] = merged_df.apply(
                    lambda x: round(x[sap_qty_col] / x["Base_Qty"], 5)
                    if x["Base_Qty"] in [1000, 100, 1] and x["Base_Qty"] != 0 else round(x[sap_qty_col], 5),
                    axis=1
                )

        if plm_qty_col:
            merged_df["PLM_Consumption"] = merged_df[plm_qty_col].fillna(0).round(5)

        # ------------------------
        # Step 5: Consumption Comparison
        # ------------------------
        merged_df["Consumption_Status"] = merged_df.apply(
            lambda x: "SAP Consumption Higher"
            if x["SAP_Consumption"] > x["PLM_Consumption"]
            else "OK",
            axis=1
        )

        merged_df["Consumption_Difference"] = (
            merged_df["SAP_Consumption"] - merged_df["PLM_Consumption"]
        ).round(5)

        # ------------------------
        # Step 6: Similarity Scores
        # ------------------------
        def safe_ratio(a, b):
            try:
                return fuzz.token_sort_ratio(str(a), str(b))
            except Exception:
                return 0

        merged_df["Material_Similarity"] = merged_df.apply(
            lambda x: safe_ratio(x.get("Material", ""), x.get("Material", "")), axis=1
        )
        merged_df["Color_Similarity"] = merged_df.apply(
            lambda x: safe_ratio(x.get("Color_SAP", ""), x.get("Color_PLM", "")), axis=1
        )
        merged_df["Consumption_Similarity"] = merged_df.apply(
            lambda x: safe_ratio(x.get("SAP_Consumption", ""), x.get("PLM_Consumption", "")), axis=1
        )

        # ------------------------
        # Step 7: Summary Counts (deduplicated)
        # ------------------------
        summary_df = merged_df.drop_duplicates(subset=["Material"])
        total_rows = len(summary_df)
        matched_materials = (summary_df["Material_Match"] == "Matched").sum()
        invalid_components = (summary_df["Component_Flag"] == "Check (Invalid)").sum()
        sap_higher = (summary_df["Consumption_Status"] == "SAP Consumption Higher").sum()

        st.subheader("📈 Summary Overview (Unique Materials)")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Unique Materials", total_rows)
        c2.metric("Material Matches", matched_materials)
        c3.metric("Invalid Components", invalid_components)
        c4.metric("SAP Higher Consumption", sap_higher)

        # ------------------------
        # Step 8: Export
        # ------------------------
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            merged_df.to_excel(writer, sheet_name="Comparison_Report", index=False)
            summary_df.to_excel(writer, sheet_name="Summary_Unique", index=False)
            if len(sap_dupes) > 0:
                sap_dupes.to_excel(writer, sheet_name="SAP_Duplicates", index=False)
            if len(plm_dupes) > 0:
                plm_dupes.to_excel(writer, sheet_name="PLM_Duplicates", index=False)
        output.seek(0)

        st.download_button(
            label="📥 Download Full Comparison Report",
            data=output,
            file_name="SAP_PLM_Comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # ------------------------
        # Step 9: Preview
        # ------------------------
        st.subheader("🔍 Preview of Comparison Results")
        preview_cols = [
            "Material", "Material Description_SAP", "Vendor Reference_SAP", "Vendor Reference_PLM",
            "Component_Flag", "Base_Qty", "SAP_Consumption", "PLM_Consumption",
            "Consumption_Difference", "Material_Match", "VendorRef_Status", "Consumption_Status",
            "Material_Similarity", "Color_Similarity", "Consumption_Similarity"
        ]
        available_cols = [c for c in preview_cols if c in merged_df.columns]
        st.dataframe(merged_df[available_cols].head(100))

    except Exception as e:
        st.error(f"❌ Error while processing: {e}")

else:
    st.info("⬆️ Please upload both SAP and PLM files to start comparison.")
