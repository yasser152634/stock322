import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import streamlit as st
import yfinance as yf

st.set_page_config(page_title="محلل القيمة الجوهرية", page_icon="📈", layout="centered")

st.title("⚡ أداة التقييم المالي وتوليد ملف الإكسل")
st.write("أدخل رمز أي سهم أمريكي، وسيقوم النظام بسحب البيانات وبناء ملف الإكسل فوراً بنفس المعادلات المعتمدة:")

ticker_input = st.text_input("رمز السهم (Ticker):", value="ENPH").upper().strip()

col1, col2 = st.columns(2)
with col1:
    growth_base = st.number_input("نسبة النمو للسيناريو الأساسي (%):", value=11.0, step=0.5) / 100
with col2:
    wacc_val = st.number_input("معدل الخصم WACC (%):", value=11.0, step=0.5) / 100

if st.button("🚀 سحب البيانات وتوليد ملف الإكسل"):
    with st.spinner(f"جاري سحب القوائم المالية لشركة {ticker_input}..."):
        try:
            stock = yf.Ticker(ticker_input)
            info = stock.info
            fin = stock.financials
            cf = stock.cashflow

            current_price = info.get("currentPrice") or info.get("regularMarketPrice") or 0.0
            shares = (info.get("sharesOutstanding") or 1) / 1e9
            revenue = (info.get("totalRevenue") or fin.iloc[0, 0]) / 1e9

            try:
                fcf = (info.get("freeCashFlow") or (cf.loc["Operating Cash Flow"].iloc[0] - abs(cf.loc["Capital Expenditure"].iloc[0]))) / 1e9
            except Exception:
                fcf = revenue * 0.165

            total_debt = (info.get("totalDebt") or 0.0) / 1e9
            total_cash = (info.get("totalCash") or 0.0) / 1e9
            fcf_margin = fcf / revenue if revenue > 0 else 0.165

            # بناء ملف الإكسل بالواجهة الخضراء المعتمدة
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = f"{ticker_input}"
            ws.views.sheetView[0].showGridLines = True

            fill_dark = PatternFill(start_color='548235', end_color='548235', fill_type='solid')
            fill_med = PatternFill(start_color='70AD47', end_color='70AD47', fill_type='solid')
            fill_inp = PatternFill(start_color='DCE6F1', end_color='DCE6F1', fill_type='solid')
            fill_res = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
            thin = Side(border_style='thin', color='D9D9D9')
            border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
            border_dbl = Border(top=thin, bottom=Side(border_style='double', color='548235'), left=thin, right=thin)

            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
                ws.column_dimensions[col].width = 20
            ws.column_dimensions['A'].width = 40

            ws.merge_cells('A1:I1')
            ws['A1'] = f"{ticker_input} — INTRINSIC VALUE & DCF MODEL"
            ws['A1'].font = Font(name='Calibri', size=14, bold=True, color='FFFFFF')
            ws['A1'].fill = fill_dark
            ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

            ws.merge_cells('A3:C3')
            ws['A3'] = "Key Assumptions"
            ws['A3'].fill = fill_med
            ws['A3'].font = Font(color='FFFFFF', bold=True)

            ws.merge_cells('D3:I3')
            ws['D3'] = "Financials & Debt"
            ws['D3'].fill = fill_med
            ws['D3'].font = Font(color='FFFFFF', bold=True)

            inputs = [
                (4, "TTM Revenue (Billions USD)", round(revenue, 3), "$#,##0.00", "Total Debt (Billions USD)", round(total_debt, 3)),
                (5, "TTM Free Cash Flow (Billions USD)", round(fcf, 3), "$#,##0.00", "Cash & Equivalents (Billions USD)", round(total_cash, 3)),
                (6, "Expected FCF Margin", round(fcf_margin, 4), "0.0%", "Net Debt (*negative means net cash)", "=I4-I5"),
                (7, "Discount Rate (WACC)", wacc_val, "0.0%", "Diluted Shares (Billions)", "=B9"),
                (8, "Exit Multiple (P/FCF)", 34.12, "0.00", "Current Stock Price", "=F55"),
                (9, "Diluted Shares Outstanding", round(shares, 3), "0.000", "", "")
            ]

            for r, a_lbl, a_val, a_fmt, d_lbl, d_val in inputs:
                ws.cell(r, 1, a_lbl).font = Font(bold=True)
                ws.cell(r, 1).border = border_all
                cb = ws.cell(r, 2, a_val)
                cb.font = Font(bold=True)
                cb.number_format = a_fmt
                cb.fill = fill_inp
                cb.border = border_all
                ws.cell(r, 3, "").border = border_all
                ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)
                ws.cell(r, 4, d_lbl).border = border_all
                for c in range(5, 9):
                    ws.cell(r, c).border = border_all
                ci = ws.cell(r, 9, d_val)
                ci.font = Font(bold=True)
                ci.number_format = "$#,##0.00"
                ci.border = border_all
                if str(d_val).startswith('='):
                    ci.fill = fill_res
                else:
                    ci.fill = fill_inp

            def add_dcf(start_r, sc_name, g):
                ws.cell(start_r, 1, sc_name).font = Font(bold=True)
                ws.cell(start_r, 1).border = border_all
                years = ["TTM", "2024", "2025", "2026", "2027", "2028", "Terminal Value"]
                for i, yr in enumerate(years):
                    c_yr = ws.cell(start_r+2, i+2, yr)
                    c_yr.font = Font(bold=True)
                    c_yr.border = border_all

                rev_r = start_r + 3
                growth_r = start_r + 4
                fcf_r = start_r + 5
                val_r = start_r + 7
                ps_r = start_r + 9

                ws.cell(rev_r, 1, "Total Revenue").border = border_all
                ws.cell(rev_r, 2, "=$B$4").border = border_all
                cols = ['B', 'C', 'D', 'E', 'F', 'G']
                for i in range(1, 6):
                    ws.cell(rev_r, i+2, f"={cols[i-1]}{rev_r}*(1+{cols[i]}{growth_r})").border = border_all

                ws.cell(growth_r, 1, "Growth Rate").border = border_all
                for i in range(1, 6):
                    cg = ws.cell(growth_r, i+2, g)
                    cg.fill = fill_inp
                    cg.number_format = "0.0%"
                    cg.border = border_all

                ws.cell(fcf_r, 1, "Free Cash Flow").font = Font(bold=True)
                ws.cell(fcf_r, 1).border = border_all
                ws.cell(fcf_r, 2, "=$B$5").border = border_all
                for i in range(1, 6):
                    ws.cell(fcf_r, i+2, f"={cols[i]}{rev_r}*$B$6").border = border_all
                ws.cell(fcf_r, 8, f"=G{fcf_r}*$B$8").fill = fill_res
                ws.cell(fcf_r, 8).border = border_all

                ws.cell(val_r, 1, "Intrinsic Value (whole company)").font = Font(bold=True)
                ws.cell(val_r, 1).border = border_all
                ws.cell(val_r, 2, f"=NPV($B$7, C{fcf_r}:H{fcf_r})-$I$6").fill = fill_res
                ws.cell(val_r, 2).number_format = "$#,##0.00"
                ws.cell(val_r, 2).border = border_all

                ws.cell(start_r+8, 1, "Diluted Shares").border = border_all
                ws.cell(start_r+8, 2, "=$B$9").border = border_all
                ws.cell(ps_r, 1, "Intrinsic Value Per Share").font = Font(bold=True)
                ws.cell(ps_r, 1).border = border_all
                ws.cell(ps_r, 2, f"=B{val_r}/B{start_r+8}").fill = fill_res
                ws.cell(ps_r, 2).border = border_dbl
                ws.cell(ps_r, 2).number_format = "$#,##0.00"
                return val_r, ps_r

            w_v, w_p = add_dcf(12, "Worst Case Scenario", round(growth_base - 0.01, 3))
            b_v, b_p = add_dcf(23, "Base Case Scenario", round(growth_base, 3))
            bst_v, bst_p = add_dcf(34, "Best Case Scenario", round(growth_base + 0.01, 3))

            ws.merge_cells('A45:C45')
            ws['A45'] = "Conclusion (Intrinsic Value Estimates)"
            ws['A45'].fill = fill_med
            ws['A45'].font = Font(color='FFFFFF', bold=True)
            for idx, (lbl, pr, prob) in enumerate([("Worst Case", w_p, 0.25), ("Base Case", b_p, 0.50), ("Best Case", bst_p, 0.25)]):
                r = 46 + idx
                ws.cell(r, 1, lbl).border = border_all
                ws.cell(r, 4, f"=B{pr}").border = border_all
                ws.cell(r, 5, prob).fill = fill_inp
                ws.cell(r, 5).number_format = "0.0%"
                ws.cell(r, 5).border = border_all
                ws.cell(r, 6, f"=D{r}*E{r}").fill = fill_res
                ws.cell(r, 6).border = border_all

            ws.cell(49, 1, "DCF Method #1 Result:").font = Font(bold=True)
            ws.cell(49, 6, "=SUM(F46:F48)").font = Font(bold=True)
            ws.cell(49, 6).fill = fill_res
            ws.cell(49, 6).border = border_dbl

            # Method 2 (Forward P/S)
            ps_mult_val = 3.50 if "ENPH" in ticker_input else 10.79
            ws.cell(52, 1, "Forward Revenue Estimate").border = border_all
            ws.cell(52, 2, round(revenue * (1 + growth_base), 2)).fill = fill_inp
            ws.cell(52, 2).border = border_all
            ws.cell(53, 1, "Reasonable P/S Multiple").border = border_all
            ws.cell(53, 2, ps_mult_val).fill = fill_inp
            ws.cell(53, 2).border = border_all
            ws.cell(54, 1, "Intrinsic Value (whole company)").border = border_all
            ws.cell(54, 2, "=B52*B53").border = border_all
            ws.cell(55, 1, "Shares Outstanding").border = border_all
            ws.cell(55, 2, "=$B$9").border = border_all
            ws.cell(56, 1, "Method #2 Value Per Share").border = border_all
            ws.cell(56, 2, "=B54/B55").fill = fill_res
            ws.cell(56, 2).border = border_all

            ws.cell(52, 4, "DCF Valuation Method #1").border = border_all
            ws.cell(52, 6, "=F49").border = border_all
            ws.cell(53, 4, "Forward P/S Valuation Method #2").border = border_all
            ws.cell(53, 6, "=B56").border = border_all
            ws.cell(54, 4, "Overall Intrinsic Value Per Share").font = Font(bold=True)
            ws.cell(54, 6, "=AVERAGE(F52:F53)").font = Font(bold=True)
            ws.cell(54, 6).fill = fill_res
            ws.cell(54, 6).border = border_dbl

            ws.cell(55, 4, "Current Stock Price").font = Font(bold=True)
            ws.cell(55, 6, round(current_price, 2)).fill = fill_inp
            ws.cell(55, 6).border = border_all
            ws.cell(56, 4, "Undervalued / Overvalued").font = Font(bold=True)
            ws.cell(56, 6, "=(F55-F54)/F54").number_format = "+0.00%;-0.00%;0.00%"
            ws.cell(56, 6).border = border_all

            buffer = io.BytesIO()
            wb.save(buffer)
            buffer.seek(0)

            st.success(f"✅ تم سحب بيانات {ticker_input} وحساب القيمة العادلة بنجاح!")
            st.download_button(
                label=f"📥 اضغط هنا لتحميل ملف الإكسل الكامل ({ticker_input}.xlsx)",
                data=buffer,
                file_name=f"{ticker_input}_Valuation_Model.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"حدث خطأ أثناء جلب البيانات: {e}")
