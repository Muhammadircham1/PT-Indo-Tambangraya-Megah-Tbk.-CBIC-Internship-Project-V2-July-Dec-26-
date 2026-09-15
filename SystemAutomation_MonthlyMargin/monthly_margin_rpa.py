import os
import glob
import win32com.client
import sys
import traceback

class BusinessRuleException(Exception):
    pass

def get_last_row(sheet, col=1):
    return sheet.Cells(sheet.Rows.Count, col).End(-4162).Row # xlUp = -4162

def get_last_col(sheet, row=1):
    return sheet.Cells(row, sheet.Columns.Count).End(-4159).Column # xlToLeft = -4159

def clear_column_data(sheet, col, start_row=2):
    lr = get_last_row(sheet, col)
    if lr >= start_row:
        sheet.Range(sheet.Cells(start_row, col), sheet.Cells(lr, col)).ClearContents()

def find_column_by_header(sheet, header_row, header_text, exact=False):
    last_col = get_last_col(sheet, header_row)
    for c in range(1, last_col + 1):
        val = sheet.Cells(header_row, c).Value
        if val is not None:
            cleaned_val = str(val).strip().replace('\n', ' ')
            # Collapse multiple spaces
            cleaned_val = ' '.join(cleaned_val.split())
            if exact:
                if header_text.lower() == cleaned_val.lower():
                    return c
            else:
                if header_text.lower() in cleaned_val.lower():
                    return c
    return None

import argparse

def main():
    try:
        parser = argparse.ArgumentParser(description="Monthly Margin RPA")
        parser.add_argument("--master", type=str, help="Path to Master Data", required=False)
        parser.add_argument("--summary", type=str, help="Path to Summary Loading", required=False)
        parser.add_argument("--profitability", type=str, help="Path to Profitability Actual", required=False)
        args = parser.parse_args()

        base_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Determine paths (use args if provided, fallback to default behavior)
        master_data_path = args.master if args.master else os.path.join(base_dir, 'MasterData.xlsx')
        if not os.path.exists(master_data_path):
            raise Exception(f"MasterData.xlsx not found at {master_data_path}")
            
        if args.summary:
            summary_loading_path = args.summary
        else:
            summary_loading_files = glob.glob(os.path.join(base_dir, "*Summary Loading of ITM*.xls*"))
            if not summary_loading_files:
                raise BusinessRuleException("INVALID FILE FORMAT: Cek kembali file yang diunduh (*Summary Loading of ITM* not found)")
            summary_loading_path = summary_loading_files[0]
            
        if args.profitability:
            profitability_path = args.profitability
        else:
            profitability_files = glob.glob(os.path.join(base_dir, "*Profitability Actual*.xls*"))
            if not profitability_files:
                raise BusinessRuleException("INVALID FILE FORMAT: Cek kembali file yang diunduh (*Profitability Actual* not found)")
            profitability_path = profitability_files[0]
            
        print(f"Master Data Path: {master_data_path}")
        print(f"Summary Loading Path: {summary_loading_path}")
        print(f"Profitability Path: {profitability_path}")

        # Initialize Excel
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False # Run in background
        excel.DisplayAlerts = False
        
        # Performance optimization
        excel.ScreenUpdating = False
        
        # Open Workbooks
        wb_master = excel.Workbooks.Open(master_data_path)
        wb_summary = excel.Workbooks.Open(summary_loading_path)
        wb_profitability = excel.Workbooks.Open(profitability_path)
        
        excel.Calculation = -4135 # xlCalculationManual
        
        # Pre-flight Check: Sheet existence
        summary_sheets = [sh.Name for sh in wb_summary.Sheets]
        profitability_sheets = [sh.Name for sh in wb_profitability.Sheets]
        
        if 'ITM Summary Actual' not in summary_sheets or 'Var' not in summary_sheets:
            raise BusinessRuleException("INVALID FILE FORMAT: Cek kembali file yang diunduh")
            
        prof_sheet_name = 'Sheet1'
        if 'Sheet1' not in profitability_sheets:
            if 'forSSO' in profitability_sheets:
                prof_sheet_name = 'forSSO'
            elif 'Update' in profitability_sheets:
                prof_sheet_name = 'Update'
            else:
                raise BusinessRuleException("INVALID FILE FORMAT: Cek kembali file yang diunduh (Sheet forSSO/Update not found)")
            
        # =========================================================================
        # PHASE 1: INGESTION DATA MENTAH DARI SUMBER 1
        # =========================================================================
        print("Phase 1: Ingestion Data Mentah...")
        
        sh_itm_actual = wb_summary.Sheets("ITM Summary Actual")
        sh_master_actual = wb_master.Sheets("Summary_Actual")
        sh_master_actual.Cells.Clear()
        sh_itm_actual.UsedRange.Copy()
        sh_master_actual.Cells(1, 1).PasteSpecial(Paste=-4163) # xlPasteValues
        
        if 'ITM Summary Baseline' in summary_sheets:
            sh_itm_baseline = wb_summary.Sheets("ITM Summary Baseline")
            sh_master_baseline = wb_master.Sheets("Summary_Baseline")
            sh_master_baseline.Cells.Clear()
            sh_itm_baseline.UsedRange.Copy()
            sh_master_baseline.Cells(1, 1).PasteSpecial(Paste=-4163)
        
        # =========================================================================
        # PHASE 2: DATA MAPPING OPERASIONAL ("Fact_Shipments")
        # =========================================================================
        print("Phase 2: Data Mapping (Fact_Shipments)...")
        sh_fact = wb_master.Sheets("Fact_Shipments")
        
        h1 = None
        for r in range(1, 10):
            val = sh_master_actual.Cells(r, 2).Value
            if val and str(val).strip() == "Shipment":
                h1 = r
                break
        
        if not h1:
            h1 = 3 # fallback
            
        def find_dollar_metric(sheet, h1, keyword):
            last_col = sheet.Cells(h1, sheet.Columns.Count).End(-4159).Column
            for c in range(1, last_col + 1):
                val1 = sheet.Cells(h1, c).Value
                if val1 and keyword in str(val1):
                    val2 = sheet.Cells(h1 + 1, c).Value
                    if val2 and str(val2).strip() == "$":
                        return c
            return None

        col_qty = find_column_by_header(sh_master_actual, h1, "Total")
        col_revenue = find_dollar_metric(sh_master_actual, h1, "Revenue")
        col_demurrage = find_dollar_metric(sh_master_actual, h1, "Demurrage")
        col_margin = find_dollar_metric(sh_master_actual, h1, "Margin")
        col_penalty = find_dollar_metric(sh_master_actual, h1, "Bonus/Penalty")
        if not col_penalty: col_penalty = find_dollar_metric(sh_master_actual, h1 + 1, "Bonus/Penalty")
        col_no = find_column_by_header(sh_master_actual, h1, "No.") or find_column_by_header(sh_master_actual, h1 + 1, "No.")
        col_month = find_column_by_header(sh_master_actual, h1, "Month") or find_column_by_header(sh_master_actual, h1 + 1, "Month")
        col_company = find_column_by_header(sh_master_actual, h1, "Company") or find_column_by_header(sh_master_actual, h1 + 1, "Company")
        col_vessel = find_column_by_header(sh_master_actual, h1, "Name of Vessel") or find_column_by_header(sh_master_actual, h1 + 1, "Name of Vessel")
        col_buyer = find_column_by_header(sh_master_actual, h1, "Buyer") or find_column_by_header(sh_master_actual, h1 + 1, "Buyer")
        col_port = find_column_by_header(sh_master_actual, h1, "Load Port") or find_column_by_header(sh_master_actual, h1 + 1, "Load Port")
        
        last_row_actual = get_last_row(sh_master_actual, col_no if col_no else (col_qty if col_qty else 1))
        data_start = h1 + 2

        tgt_qty = find_column_by_header(sh_fact, 2, "Quantity MT")
        tgt_rev = find_column_by_header(sh_fact, 2, "Revenue USD") or find_column_by_header(sh_fact, 2, "Revenue")
        tgt_dem = find_column_by_header(sh_fact, 2, "Demurrage USD") or find_column_by_header(sh_fact, 2, "Demurrage")
        tgt_mar = find_column_by_header(sh_fact, 2, "Margin Uplift")
        tgt_pen = find_column_by_header(sh_fact, 2, "Bonus Penalty USD") or find_column_by_header(sh_fact, 2, "Penalty")
        tgt_no = find_column_by_header(sh_fact, 2, "No")
        tgt_month = find_column_by_header(sh_fact, 2, "Month Num")
        tgt_company = find_column_by_header(sh_fact, 2, "Company")
        tgt_vessel = find_column_by_header(sh_fact, 2, "Vessel Name")
        tgt_buyer = find_column_by_header(sh_fact, 2, "Shipment")
        tgt_port = find_column_by_header(sh_fact, 2, "Load_Port")

        def copy_metric(src_col, tgt_col):
            if src_col and tgt_col:
                clear_column_data(sh_fact, tgt_col, 3)
                data = sh_master_actual.Range(sh_master_actual.Cells(data_start, src_col), sh_master_actual.Cells(last_row_actual, src_col)).Value
                sh_fact.Range(sh_fact.Cells(3, tgt_col), sh_fact.Cells(last_row_actual - data_start + 3, tgt_col)).Value = data

        copy_metric(col_no, tgt_no)
        copy_metric(col_month, tgt_month)
        copy_metric(col_company, tgt_company)
        copy_metric(col_vessel, tgt_vessel)
        copy_metric(col_buyer, tgt_buyer)
        copy_metric(col_port, tgt_port)
        copy_metric(col_qty, tgt_qty)
        copy_metric(col_revenue, tgt_rev)
        copy_metric(col_demurrage, tgt_dem)
        copy_metric(col_margin, tgt_mar)
        copy_metric(col_penalty, tgt_pen)

        sh_deviasi = wb_summary.Sheets("Deviasi")
        
        h_dev = None
        for r in range(1, 10):
            val = sh_deviasi.Cells(r, 2).Value
            if val and str(val).strip() == "Shipment":
                h_dev = r
                break
        if not h_dev:
            h_dev = 4
            
        col_sso = find_column_by_header(sh_deviasi, h_dev, "Implementation")
        if col_sso:
            last_row_sso = get_last_row(sh_deviasi, col_sso)
            data_start_sso = h_dev + 2
            
            tgt_sso_flag = find_column_by_header(sh_fact, 2, "SSO_Flag")
            if tgt_sso_flag:
                clear_column_data(sh_fact, tgt_sso_flag, 3)
                sh_deviasi.Range(sh_deviasi.Cells(data_start_sso, col_sso), sh_deviasi.Cells(last_row_sso, col_sso)).Copy()
                sh_fact.Cells(3, tgt_sso_flag).PasteSpecial(Paste=-4163)
                
            tgt_sso_pct = find_column_by_header(sh_fact, 2, "SSO_Presentase")
            if tgt_sso_pct:
                clear_column_data(sh_fact, tgt_sso_pct, 3)
                sh_deviasi.Range(sh_deviasi.Cells(data_start_sso, col_sso), sh_deviasi.Cells(last_row_sso, col_sso)).Copy()
                sh_fact.Cells(3, tgt_sso_pct).PasteSpecial(Paste=-4163) 
        
        # Use the same col_no logic to accurately find the last valid shipment row
        col_no_actual = find_column_by_header(sh_master_actual, h1, "No.") or find_column_by_header(sh_master_actual, h1, "No") or find_column_by_header(sh_master_actual, h1 + 1, "No.") or find_column_by_header(sh_master_actual, h1 + 1, "No")
        last_row_actual = get_last_row(sh_master_actual, col_no_actual if col_no_actual else (col_qty if col_qty else 1))
        tgt_last_row = last_row_actual - data_start + 3
        
        # Clear old rows beyond tgt_last_row to prevent old formulas from inflating row count and to remove leftover borders
        if tgt_last_row >= 3:
            try:
                sh_fact.Rows(str(tgt_last_row + 1) + ":" + str(sh_fact.Rows.Count)).Delete()
            except:
                sh_fact.Rows(str(tgt_last_row + 1) + ":" + str(sh_fact.Rows.Count)).Clear()
            
        import re
        last_col_fact = get_last_col(sh_fact, 2)
        print(f"Fact_Shipments last_col_fact: {last_col_fact}, tgt_last_row: {tgt_last_row}")
        for c in range(1, last_col_fact + 1):
            if sh_fact.Cells(3, c).HasFormula:
                fmla = str(sh_fact.Cells(3, c).Formula)
                print(f"Col {c} has formula: {fmla}")
                if ":" in fmla and re.search(r'([A-Z]+)\$?3:[A-Z]+\$?\d+', fmla):
                    new_fmla = re.sub(r'([A-Z]+)\$?3:[A-Z]+\$?\d+', r'\g<1>3', fmla)
                    print(f"Clearing anchor and replacing array formula: {fmla} -> {new_fmla}")
                    sh_fact.Cells(3, c).ClearContents()
                    sh_fact.Cells(3, c).Formula = new_fmla
                print(f"AutoFilling Col {c} down to row {tgt_last_row}")
                sh_fact.Range(sh_fact.Cells(3, c), sh_fact.Cells(3, c)).AutoFill(sh_fact.Range(sh_fact.Cells(3, c), sh_fact.Cells(tgt_last_row, c)))

        # Force fill for hardcoded Year (Col 2) and Month Name (Col 4) using Copy/Paste to avoid date incrementing
        if tgt_last_row >= 4:
            sh_fact.Range(sh_fact.Cells(3, 2), sh_fact.Cells(3, 2)).Copy()
            sh_fact.Range(sh_fact.Cells(4, 2), sh_fact.Cells(tgt_last_row, 2)).PasteSpecial(Paste=-4163)
            sh_fact.Range(sh_fact.Cells(3, 4), sh_fact.Cells(3, 4)).Copy()
            sh_fact.Range(sh_fact.Cells(4, 4), sh_fact.Cells(tgt_last_row, 4)).PasteSpecial(Paste=-4163)

        # Format borders dynamically to follow the number of shipments by copying format from row 3
        if tgt_last_row >= 4:
            sh_fact.Range(sh_fact.Cells(3, 1), sh_fact.Cells(3, last_col_fact)).Copy()
            sh_fact.Range(sh_fact.Cells(4, 1), sh_fact.Cells(tgt_last_row, last_col_fact)).PasteSpecial(Paste=-4122) # xlPasteFormats


        # =========================================================================
        # PHASE 3: KALKULASI INTERNAL MATRIKS
        # =========================================================================
        print("Phase 3: Kalkulasi Internal Matriks...")
        
        tgt_qty = find_column_by_header(sh_fact, 2, "Quantity MT")
        true_last_row = get_last_row(sh_fact, tgt_qty if tgt_qty else 9) # Use Quantity MT column
        
        sh_var = wb_summary.Sheets("Var")
        # Column 1 is empty, use column 2 (No) for last row detection in Var sheet
        # Wait, Col D is the dynamic array that spills. Let's get last row from Col D.
        last_row_var = get_last_row(sh_var, 4) # End User (dynamic array)
        
        # AutoFill formulas in Var sheet down to last_row_var
        last_col_var = get_last_col(sh_var, 3)
        for c in range(2, last_col_var + 1):
            if c == 4: continue # Skip End User since it's a dynamic array
            if sh_var.Cells(4, c).HasFormula:
                sh_var.Range(sh_var.Cells(4, c), sh_var.Cells(4, c)).AutoFill(sh_var.Range(sh_var.Cells(4, c), sh_var.Cells(last_row_var, c)))
        
        # Force calculation so AutoFilled formulas evaluate before we copy them!
        excel.Calculate()
        
        # Sheet 1
        sh1 = wb_master.Sheets("Sheet1")
        last_col_sh1 = get_last_col(sh1, 1)
        import re
        if last_col_sh1 > 1:
             # Sheet1 is a summary table (rows 2 to 4), do not autofill down to true_last_row
             for r in range(2, 5):
                 for c in range(1, last_col_sh1 + 1):
                     if sh1.Cells(r, c).HasFormula:
                         old_f = sh1.Cells(r, c).Formula
                         # Replace Fact_Shipments ranges safely to avoid circular references
                         new_f = re.sub(r'(Fact_Shipments![A-Z\$]+\d+:[A-Z\$]+)\d+', r'\g<1>' + str(true_last_row), old_f)
                         if old_f != new_f:
                             sh1.Cells(r, c).Formula = new_f
                         
                         # Apply 2 decimal places rounding for 'Implementation' and 'Non-Implementation' visually
                         header_val = str(sh1.Cells(1, c).Value).strip()
                         if header_val in ["Implementation", "Non-Implementation"]:
                             sh1.Cells(r, c).NumberFormat = "0.00"
                     
        col_dem_ton = find_column_by_header(sh1, 1, "Dem $/ton")
        if col_dem_ton:
            # Copy value only for the summary rows (2 to 4)
            rng = sh1.Range(sh1.Cells(2, col_dem_ton), sh1.Cells(4, col_dem_ton))
            rng.Copy()
            rng.PasteSpecial(Paste=-4163)
            
        col_denda = find_column_by_header(sh1, 1, "Denda/Demurrage")
        if not col_denda:
            col_denda = find_column_by_header(sh1, 1, "Demurrage Value")
            
        if col_denda:
            for r in range(2, 5):
                dem_ton_val = sh1.Cells(r, col_dem_ton).Value if col_dem_ton else 0
                qty_val = sh1.Cells(r, 4).Value # Throughput
                if isinstance(dem_ton_val, (int, float)) and isinstance(qty_val, (int, float)):
                    sh1.Cells(r, col_denda).Value = (dem_ton_val * -1) * qty_val

        # Sheet 2
        sh2 = wb_master.Sheets("Sheet2")
        last_col_sh2 = get_last_col(sh2, 2)
        if last_col_sh2 > 1:
             for c in range(1, last_col_sh2 + 1):
                 if sh2.Cells(3, c).HasFormula:
                     clear_column_data(sh2, c, 3)
             if last_row_var >= 3:
                 # Clear garbage rows in Sheet2
                 sh2.Rows(str(last_row_var + 1) + ":" + str(sh2.Rows.Count)).ClearContents()
                 for c in range(1, last_col_sh2 + 1):
                     if sh2.Cells(3, c).HasFormula:
                         sh2.Range(sh2.Cells(3, c), sh2.Cells(3, c)).AutoFill(sh2.Range(sh2.Cells(3, c), sh2.Cells(last_row_var, c)))
                 
                 # Force calculation for Sheet2 as well
                 excel.Calculate()

        cols_var = {
            "No": "No",
            "Company": "Company",
            "End User": "End User",
            "Status": "Status",
            "Total Baseline": "TP Plan (MT)",
            "Total Actual": "TP Actual\n(MT)",
            "Margin Baseline": "Margin Plan\n($)",
            "Margin Actual": "Margin Actual\n($)"
        }
        
        for k_var, k_sh2 in cols_var.items():
            var_col = find_column_by_header(sh_var, 3, k_var)
            if var_col:
                tgt_col = find_column_by_header(sh2, 2, k_sh2)
                # Fallback to fuzzy match if exact match fails
                if not tgt_col:
                    if "TP Plan" in k_sh2: tgt_col = 5
                    elif "TP Actual" in k_sh2: tgt_col = 6
                    elif "Margin Plan" in k_sh2: tgt_col = 7
                    elif "Margin Actual" in k_sh2: tgt_col = 8
                    
                if tgt_col:
                    print(f"Copying {k_var} from Var col {var_col} to Sheet2 col {tgt_col} (rows 4 to {last_row_var})")
                    clear_column_data(sh2, tgt_col, 3)
                    if last_row_var >= 4:
                        sh_var.Range(sh_var.Cells(4, var_col), sh_var.Cells(last_row_var, var_col)).Copy()
                        sh2.Cells(3, tgt_col).PasteSpecial(Paste=-4163)
                        print(f"Verified row 49 of col {tgt_col}: {sh2.Cells(49, tgt_col).Value}")
        
        # Sheet 3
        sh3 = wb_master.Sheets("Sheet3")
        last_col_sh3 = get_last_col(sh3, 2)
        if last_col_sh3 > 1:
             # Sheet3 is a summary table, do not autofill. Just update formulas to point to last_row_var
             for r in range(2, 7): # rows 2, 3, 4, 5, 6
                 for c in range(1, last_col_sh3 + 1):
                     if sh3.Cells(r, c).HasFormula:
                         old_f = sh3.Cells(r, c).Formula
                         # Replace hardcoded row limits safely to avoid circular references
                         new_f = re.sub(r'(Sheet2![A-Z\$]+\d+:[A-Z\$]+)\d+', r'\g<1>' + str(last_row_var), old_f)
                         if old_f != new_f:
                             sh3.Cells(r, c).Formula = new_f

        # =========================================================================
        # PHASE 4: INTEGRASI DATA PROFITABILITAS (SUMBER 2)
        # =========================================================================
        print(f"Phase 4: Integrasi Data Profitabilitas (using {prof_sheet_name})...")
        sh_prof = wb_profitability.Sheets(prof_sheet_name)
        sh4 = wb_master.Sheets("Sheet4")
        
        # Clear all old data from row 6 down to avoid lingering rows
        last_r_sh4 = get_last_row(sh4, 1)
        if last_r_sh4 >= 6:
             sh4.Rows("6:" + str(sh4.Rows.Count)).ClearContents()
        
        # Find header row and company column dynamically in Profitability
        header_row_prof = 3
        col_company = 1
        for r in range(1, 10):
            for c in range(1, 10):
                val = sh_prof.Cells(r, c).Value
                if val and str(val).strip() == "Company":
                    header_row_prof = r
                    col_company = c
                    break
                    
        last_col_prof = get_last_col(sh_prof, header_row_prof)
        last_row_prof = get_last_row(sh_prof, col_company)
        
        # Mapping from Profitability (forSSO) to MasterData (Sheet4)
        target_mapping = {
            "Company": ["Company"],
            "Customer": ["Customer"],
            "FOB Price Actual ($/ton)": ["FOBPrice", "FOB Price", "Sum of FOBPriceAct($/ton)"],
            "Margin Cash Actual ($/ton)": ["Margin Cash", "Sum of CashMargin($/ton)"],
            "Revenue Actual ($)": ["Sum of Revenue", "Revenue"],
            "Total Cost SSO Actual ($)": ["Sum of TotalCostforSSO$", "TotalCostforSSO$"],
            "Cash Margin Actual ($)": ["CashMargin ($)", "Cash Margin ($)"]
        }
        
        for c in range(1, last_col_prof + 1):
            val = sh_prof.Cells(header_row_prof, c).Value
            if not val: continue
            
            header = str(val).strip().replace('\n', ' ')
            header = ' '.join(header.split())
            
            tgt_header_name = None
            for tgt, aliases in target_mapping.items():
                if header in aliases:
                    tgt_header_name = tgt
                    break
                    
            if not tgt_header_name:
                continue
            
            tgt_col = find_column_by_header(sh4, 5, tgt_header_name, exact=True) # Headers in Sheet4 are on row 5
            if tgt_col:
                # Copy data from just below header downwards
                sh_prof.Range(sh_prof.Cells(header_row_prof + 1, c), sh_prof.Cells(last_row_prof, c)).Copy()
                # Paste to row 6 in Sheet4
                sh4.Cells(6, tgt_col).PasteSpecial(Paste=-4163)
        
        sh5 = wb_master.Sheets("Sheet5")
        last_row_sh4 = get_last_row(sh4, 1)
        last_col_sh4 = get_last_col(sh4, 5)
        
        found_row = None
        for r in range(last_row_sh4, 0, -1):
            for c in range(1, last_col_sh4 + 1):
                val = str(sh4.Cells(r, c).Value).strip().lower()
                if "grand total" in val:
                    found_row = r
                    break
            if found_row:
                break
                
        if found_row:
            # Clear existing data in Sheet5 starting row 3
            sh5.Rows("3:" + str(sh5.Rows.Count)).ClearContents()
            
            last_col_sh5 = get_last_col(sh5, 2)
            
            for c_sh5 in range(1, last_col_sh5 + 1):
                header_sh5 = str(sh5.Cells(2, c_sh5).Value).strip().lower()
                if not header_sh5 or header_sh5 == "none": continue
                
                for c_sh4 in range(1, last_col_sh4 + 1):
                    header_sh4 = str(sh4.Cells(5, c_sh4).Value).strip().lower()
                    if header_sh4 == header_sh5:
                        val = sh4.Cells(found_row, c_sh4).Value
                        sh5.Cells(3, c_sh5).Value = val
                        break

        print("[SUCCESS] Master Data Monthly Performance Processed.")
        
    except BusinessRuleException as be:
        print(f"BUSINESS RULE EXCEPTION: {be}")
    except Exception as e:
        print(f"SYSTEM ERROR: {e}")
        traceback.print_exc()
    finally:
        try:
            excel.Calculation = -4105 # xlCalculationAutomatic
            wb_master.Save()
            wb_master.Close()
            wb_summary.Close(False)
            wb_profitability.Close(False)
            excel.Quit()
        except:
            pass

if __name__ == "__main__":
    main()
