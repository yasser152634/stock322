import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import streamlit as st
import yfinance as yf

st.set_page_config(page_title="محلل القيمة الجوهرية والقطاع", page_icon="📊", layout="centered")

st.title("⚡ أداة التقييم المالي والتحليل القطاعي الآلي")
st.write("أدخل رمز أي سهم أمريكي، وسيقوم النظام بتحليل الشركة وأقرانها في القطاع وتوليد ملف إكسل بصفحة واحدة حية:")

ticker_input = st.text_input("رمز السهم (Ticker):", value="MSFT").upper().strip()

# خريطة مرجعية لأبرز أقران القطاعات في الخلفية
SECTOR_PEERS = {
    "SOLAR": ["ENPH", "SEDG", "FSLR", "RUN", "CSIQ"],
    "TECH_SOFTWARE": ["MSFT", "ORCL", "ADBE", "CRM", "SAP"],
    "TECH_HARDWARE": ["AAPL", "DELL", "HPQ", "LOGI"],
    "SEMICONDUCTORS": ["NVDA", "AMD", "INTC", "TSM", "AVGO", "QCOM"],
    "AUTO_EV": ["TSLA", "RIVN", "LCID", "GM", "F"],
    "ECOMMERCE": ["AMZN", "BABA", "EBAY", "WMT", "TGT"],
    "INTERNET_COMM": ["GOOGL", "META", "NFLX", "DIS"],
}

def get_peers_for_ticker(ticker, info):
    """تحديد منافسي القطاع تلقائياً بناءً على تصنيف الشركة"""
    for sector_name, group in SECTOR_PEERS.items():
        if ticker in group:
            return [p for p in group if p != ticker][:4]
    industry = (info.get("industry") or "").lower()
    sector = (info.get("sector") or "").lower()
    if "solar" in industry or "renewable" in industry:
        return ["SEDG", "FSLR", "CSIQ"]
    elif "semiconductor" in industry:
        return ["AMD", "INTC", "QCOM"]
    elif "software" in industry:
        return ["ORCL", "CRM", "ADBE"]
    elif "auto" in industry or "electric" in industry:
        return ["GM", "F", "RIVN"]
    else:
        return ["SPY"]

if st.button("🚀 تحليل الشركة والأقران وتوليد الملف"):
    with st.spinner(f"جاري سحب القوائم ومقارنة أقران القطاع لسهم {ticker_input}..."):
        try:
            stock = yf.Ticker(ticker_input)
            info = stock.info
            fin = stock.financials
            cf = stock.cashflow

            current_price = info.get("currentPrice") or info.get("regularMarketPrice") or 0.0
            shares = (info.get("sharesOutstanding") or 1) / 1e9
            revenue = (info.get("totalRevenue") or (fin.iloc[0, 0] if not fin.empty else 1.0)) / 1e9

            # حساب التدفق النقدي الحر TTM
            try:
                fcf = (info.get("freeCashFlow") or (cf.loc["Operating Cash Flow"].iloc[0] - abs(cf.loc["Capital Expenditure"].iloc[0]))) / 1e9
            except Exception:
                fcf = revenue * 0.15

            total_debt = (info.get("totalDebt") or 0.0) / 1e9
            total_cash = (info.get("totalCash") or 0.0) / 1e9
            company_fcf_margin = fcf / revenue if revenue > 0 else 0.15

            # ========================================================
            # تحليل أقران القطاع في الخلفية (Backend Peer Intelligence)
            # ========================================================
            peers = get_peers_for_ticker(ticker_input, info)
            peer_margins = [company_fcf_margin]
            peer_ps_ratios = []

            for p_sym in peers:
                try:
                    p_info = yf.Ticker(p_sym).info
                    p_rev = p_info.get("totalRevenue") or 0
                    p_fcf = p_info.get("freeCashFlow") or 0
                    if p_rev > 0 and p_fcf != 0:
                        peer_margins.append(p_fcf / p_rev)
                    p_ps = p_info.get("priceToSalesTrailing12Months")
                    if p_ps and p_ps > 0:
                        peer_ps_ratios.append(p_ps)
                except Exception:
                    continue

            # وسيط هامش التدفق النقدي للقطاع (Median)
            peer_margins.sort()
            benchmark_fcf_margin = round(peer_margins[len(peer_margins)//2], 4)

            # وسيط مكرر المبيعات للقطاع (Median)
            if peer_ps_ratios:
                peer_ps_ratios.sort()
                benchmark_ps = round(peer_ps_ratios[len(peer_ps_ratios)//2] * 0.90, 2)
            else:
                benchmark_ps = round((info.get("priceToSalesTrailing12Months") or 3.50) * 0.85, 2)

            # تكلفة رأس المال WACC الواقعية وفق نموذج CAPM
            beta = info.get("beta") or 1.15
            risk_free_rate = 0.042
            equity_risk_premium = 0.055
            dynamic_wacc = round(min(max(risk_free_rate + (beta * equity_risk_premium), 0.085), 0.145), 3)

            # مضاعف التدفق النهائي المخصص للشركة والقطاع
            trailing_pe = info.get("trailingPE") or 25.0
            exit_multiple = round(min(max(trailing_pe * 0.85, 14.0), 38.0), 2)

            # حساب معدلات النمو الواقعية بدلاً من النسب الثابتة
            rev_growth_raw = info.get("revenueGrowth")
            earn_growth_raw = info.get("earningsGrowth")
            if rev_growth_raw and rev_growth_raw > 0:
                base_growth = min(max(rev_growth_raw, 0.05), 0.28)
            elif earn_growth_raw and earn_growth_raw > 0:
                base_growth = min(max(earn_growth_raw * 0.65, 0.05), 0.22)
            else:
                base_growth = 0.10

            worst_growth = round(max(base_growth - 0.03, 0.03), 3)
            best_growth = round(base_growth + 0.03, 3)

            st.success(f"✅ تم سحب بيانات {ticker_input} وحساب مؤشرات القطاع والأقران: {', '.join(peers)}")
            st.caption(f"الوسيط القطاعي لهامش الكاش: {benchmark_fcf_margin*100:.1f}% | وسيط مكرر المبيعات: {benchmark_ps}x | معدل الخصم WACC: {dynamic_wacc*100:.1f}%")

            # ========================================================
            # بناء ملف إكسل بصفحة واحدة فقط (Single Sheet)
            # ========================================================
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = f"{ticker_input} Valuation"
            ws.views.sheetView[0].showGridLines = True

            # الألوان المعتمدة (الواجهة الخضراء)
            GREEN_DARK = '548235'
            GREEN_MED = '70AD47'
            GREEN_LIGHT = 'E2EFDA'
            INPUT_BLUE = 'DCE6F1'
            BORDER_GRAY = 'D9D9D9'

            fill_title = PatternFill(start_color=GREEN_DARK, end_color=GREEN_DARK, fill_type='solid')
            fill_sec = PatternFill(start_color=GREEN_MED, end_color=GREEN_MED, fill_type='solid')
            fill_inp = PatternFill(start_color=INPUT_BLUE, end_color=INPUT_BLUE, fill_type='solid')
            fill_res = PatternFill(start_color=GREEN_LIGHT, end_color=GREEN_LIGHT, fill_type='solid')

            thin = Side(border_style='thin', color=BORDER_GRAY)
            border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
            border_dbl = Border(top=thin, bottom=Side(border_style='double', color=GREEN_DARK), left=thin, right=thin)

            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
                ws.column_dimensions[col].width = 20
            ws.column_dimensions['A'].width = 44

            # العنوان وتنبيه التحديث اللحظي
            ws.merge_cells('A1:I1')
            ws['A1'] = f"{ticker_input} — INTRINSIC VALUE & LIVE SECTOR VALUATION"
            ws['A1'].font = Font(name='Calibri', size=13, bold=True, color='FFFFFF')
            ws['A1'].fill = fill_title
            ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

            ws.merge_cells('A2:I2')
            ws['A2'] = "لتحديث سعر السهم لحظياً: اضغط Ctrl + Alt + F5 (أو من قائمة بيانات اضغط تحديث الكل)"
            ws['A2'].font = Font(name='Calibri', size=9, bold=True, color='805500')
            ws['A2'].fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
            ws['A2'].alignment = Alignment(horizontal='center', vertical='center')

            ws.merge_cells('A3:C3')
            ws['A3'] = "Key Assumptions (مؤشرات معتمدة من وسيط القطاع)"
            ws['A3'].fill = fill_sec
            ws['A3'].font = Font(color='FFFFFF', bold=True)

            ws.merge_cells('D3:I3')
            ws['D3'] = "Financials & Debt (البيانات المالية والديون بالمليار)"
            ws['D3'].fill = fill_sec
            ws['D3'].font = Font(color='FFFFFF', bold=True)

            # الافتراضات المحسوبة من القطاع
            inputs = [
                (4, "TTM Revenue (Billions USD)", round(revenue, 3), "$#,##0.00", "Total Debt (Billions USD)", round(total_debt, 3)),
                (5, "TTM Free Cash Flow (Billions USD)", round(fcf, 3), "$#,##0.00", "Cash & Equivalents (Billions USD)", round(total_cash, 3)),
                (6, "Expected FCF Margin (Sector Benchmark)", benchmark_fcf_margin, "0.0%", "Net Debt (*negative means net cash)", "=I4-I5"),
                (7, "Discount Rate (Company WACC / CAPM)", dynamic_wacc, "0.0%", "Diluted Shares (Billions)", "=B9"),
                (8, "Exit Multiple (Sector Exit P/FCF)", exit_multiple, "0.00", "Current Stock Price (Live)", "=F55"),
                (9, "Diluted Shares Outstanding", round(shares, 3), "0.000", "", "")
            ]

            for r, a_lbl, a_val, a_fmt, d_lbl, d_val in inputs:
                ws.cell(r, 1, a_lbl).font = Font(bold=True); ws.cell(r, 1).border = border_all
                cb = ws.cell(r, 2, a_val); cb.font = Font(bold=True); cb.number_format = a_fmt; cb.border = border_all
                if r in [6, 7, 8]:
                    cb.fill = fill_res
                else:
                    cb.fill = fill_inp
                ws.cell(r, 3, "").border = border_all

                ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)
                ws.cell(r, 4, d_lbl).border = border_all
                for c in range(5, 9):
                    ws.cell(r, c).border = border_all
                ci = ws.cell(r, 9, d_val); ci.font = Font(bold=True); ci.border = border_all
                if str(d_val).startswith('='):
                    ci.fill = fill_res; ci.number_format = "$#,##0.00"
                else:
                    ci.fill = fill_inp; ci.number_format = "$#,##0.00"

            # بناء سيناريوهات الـ DCF بنفس المعادلات
            def add_dcf(start_r, sc_name, g):
                ws.cell(start_r, 1, sc_name).font = Font(bold=True); ws.cell(start_r, 1).border = border_all
                years = ["TTM", "2024", "2025", "2026", "2027", "2028", "Terminal Value"]
                for i, yr in enumerate(years):
                    c_yr = ws.cell(start_r+2, i+2, yr); c_yr.font = Font(bold=True); c_yr.border = border_all

                rev_r = start_r + 3
                growth_r = start_r + 4
                fcf_r = start_r + 5
                val_r = start_r + 7
                ps_r = start_r + 9

                ws.cell(rev_r, 1, "Total Revenue").border = border_all; ws.cell(rev_r, 2, "=$B$4").border = border_all
                cols = ['B', 'C', 'D', 'E', 'F', 'G']
                for i in range(1, 6):
                    ws.cell(rev_r, i+2, f"={cols[i-1]}{rev_r}*(1+{cols[i]}{growth_r})").border = border_all

                ws.cell(growth_r, 1, "Growth Rate").border = border_all
                for i in range(1, 6):
                    cg = ws.cell(growth_r, i+2, g)
                    cg.fill = fill_inp; cg.number_format = "0.0%"; cg.border = border_all

                ws.cell(fcf_r, 1, "Free Cash Flow").font = Font(bold=True); ws.cell(fcf_r, 1).border = border_all
                ws.cell(fcf_r, 2, "=$B$5").border = border_all
                for i in range(1, 6):
                    ws.cell(fcf_r, i+2, f"={cols[i]}{rev_r}*$B$6").border = border_all
                ws.cell(fcf_r, 8, f"=G{fcf_r}*$B$8").fill = fill_res; ws.cell(fcf_r, 8).border = border_all

                # المعادلة الأصلية
                ws.cell(val_r, 1, "Intrinsic Value (whole company)").font = Font(bold=True); ws.cell(val_r, 1).border = border_all
                ws.cell(val_r, 2, f"=NPV($B$7, C{fcf_r}:H{fcf_r})-$I$6").fill = fill_res
                ws.cell(val_r, 2).number_format = "$#,##0.00"; ws.cell(val_r, 2).border = border_all

                ws.cell(start_r+8, 1, "Diluted Shares").border = border_all; ws.cell(start_r+8, 2, "=$B$9").border = border_all
                ws.cell(ps_r, 1, "Intrinsic Value Per Share").font = Font(bold=True); ws.cell(ps_r, 1).border = border_all
                ws.cell(ps_r, 2, f"=B{val_r}/B{start_r+8}").fill = fill_res; ws.cell(ps_r, 2).border = border_dbl; ws.cell(ps_r, 2).number_format = "$#,##0.00"
                return val_r, ps_r

            w_v, w_p = add_dcf(12, f"Worst Case Scenario ({worst_growth*100:.1f}%)", worst_growth)
            b_v, b_p = add_dcf(23, f"Base Case Scenario ({base_growth*100:.1f}%)", base_growth)
            bst_v, bst_p = add_dcf(34, f"Best Case Scenario ({best_growth*100:.1f}%)", best_growth)

            # ملخص النتائج (Conclusion)
            ws.merge_cells('A45:C45'); ws['A45'] = "Conclusion (Intrinsic Value Estimates)"; ws['A45'].fill = fill_sec; ws['A45'].font = Font(color='FFFFFF', bold=True)
            for idx, (lbl, pr, prob) in enumerate([("Worst Case", w_p, 0.25), ("Base Case", b_p, 0.50), ("Best Case", bst_p, 0.25)]):
                r = 46 + idx
                ws.cell(r, 1, lbl).border = border_all
                ws.cell(r, 4, f"=B{pr}").border = border_all
                ws.cell(r, 5, prob).fill = fill_inp; ws.cell(r, 5).number_format = "0.0%"; ws.cell(r, 5).border = border_all
                ws.cell(r, 6, f"=D{r}*E{r}").fill = fill_res; ws.cell(r, 6).border = border_all

            ws.cell(49, 1, "DCF Method #1 Result:").font = Font(bold=True)
            ws.cell(49, 6, "=SUM(F46:F48)").font = Font(bold=True); ws.cell(49, 6).fill = fill_res; ws.cell(49, 6).border = border_dbl

            # Method 2 (Forward P/S)
            ws.cell(52, 1, "Forward Revenue Estimate").border = border_all
            ws.cell(52, 2, round(revenue * (1 + base_growth), 2)).fill = fill_inp; ws.cell(52, 2).border = border_all
            ws.cell(53, 1, "Sector Reasonable P/S Multiple").border = border_all
            ws.cell(53, 2, benchmark_ps).fill = fill_res; ws.cell(53, 2).border = border_all
            ws.cell(54, 1, "Method #2 Intrinsic Value (whole company)").border = border_all
            ws.cell(54, 2, "=B52*B53").border = border_all
            ws.cell(55, 1, "Shares Outstanding").border = border_all
            ws.cell(55, 2, "=$B$9").border = border_all
            ws.cell(56, 1, "Method #2 Value Per Share").border = border_all
            ws.cell(56, 2, "=B54/B55").fill = fill_res; ws.cell(56, 2).border = border_all

            # الملخص النهائي مع ربط السعر اللحظي الحي
            ws.cell(52, 4, "DCF Valuation Method #1").border = border_all; ws.cell(52, 6, "=F49").border = border_all
            ws.cell(53, 4, "Forward P/S Valuation Method #2").border = border_all; ws.cell(53, 6, "=B56").border = border_all
            ws.cell(54, 4, "Overall Intrinsic Value Per Share").font = Font(bold=True)
            ws.cell(54, 6, "=AVERAGE(F52:F53)").font = Font(bold=True); ws.cell(54, 6).fill = fill_res; ws.cell(54, 6).border = border_dbl

            # رمز السهم المباشر في D55 ومعادلة السعر الحي في F55
            ws.cell(55, 4, f"{ticker_input}").font = Font(bold=True, color='002060')
            ws.cell(55, 4).fill = fill_inp; ws.cell(55, 4).alignment = Alignment(horizontal='center'); ws.cell(55, 4).border = border_all
            ws.cell(55, 5, "Current Price:").font = Font(bold=True); ws.cell(55, 5).border = border_all

            # معادلة السعر الحي: تسحب من بطاقة السهم التفاعلية أو الإغلاق الأخير
            ws.cell(55, 6, f'=IFERROR(D55.Price, {round(current_price, 2)})').font = Font(bold=True)
            ws.cell(55, 6).number_format = "$#,##0.00"; ws.cell(55, 6).fill = fill_inp; ws.cell(55, 6).border = border_all

            # نسبة التقييم تتحدث فورياً مع السعر
            ws.cell(56, 4, "Undervalued / Overvalued").font = Font(bold=True); ws.cell(56, 4).border = border_all
            ws.cell(56, 6, "=(F55-F54)/F54").font = Font(bold=True); ws.cell(56, 6).number_format = "+0.00%;-0.00%;0.00%"; ws.cell(56, 6).border = border_all

            # مستويات هامش الأمان (Margin of Safety) مربوطة بالمعادلات
            ws.merge_cells('D59:E59'); ws['D59'] = "Margin of Safety %"; ws['D59'].fill = fill_sec; ws['D59'].font = Font(color='FFFFFF', bold=True)
            ws.merge_cells('F59:G59'); ws['F59'] = "Margin of Safety Price"; ws['F59'].fill = fill_sec; ws['F59'].font = Font(color='FFFFFF', bold=True)
            for idx, pct in enumerate([0.00, 0.10, 0.20, 0.30]):
                r = 60 + idx
                ws.cell(r, 4, pct).number_format = "0%"; ws.cell(r, 4).border = border_all
                ws.cell(r, 5, f"{int(pct*100)}%").border = border_all
                ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
                ws.cell(r, 6, f"=$F$54*(1-D{r})").number_format = "$#,##0.00"; ws.cell(r, 6).fill = fill_res; ws.cell(r, 6).border = border_all

            buffer = io.BytesIO()
            wb.save(buffer)
            buffer.seek(0)

            st.download_button(
                label=f"📥 اضغط هنا لتحميل نموذج {ticker_input} المتكامل المحدّث",
                data=buffer,
                file_name=f"{ticker_input}_Valuation_Model.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"حدث خطأ أثناء جلب البيانات: {e}")
