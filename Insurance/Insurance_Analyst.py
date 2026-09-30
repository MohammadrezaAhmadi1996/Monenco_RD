import re
import pandas as pd
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from time import time


t1 = time()
FILE   = "گزارش خسارت درمان بیمه البرز- 1405-1404.xlsx"
COST   = 'هزينه پاراکلينيکي پرداخت شده'
DIS    = 'نوع بيماري'
ID_ORG = 'شناسه یکتای پرسنلی'
ID_PER = 'شناسه یکتای کدملی'
DATE   = 'تاريخ بيماري'
MON    = 'ماه بیماری'

df = pd.read_excel(FILE)
df[COST] = pd.to_numeric(df[COST], errors='coerce').fillna(0)
df[DIS]  = df[DIS].fillna('نامشخص')
df[MON] = (df[DATE].astype(str).str.strip()
           .str.replace('-', '/', regex=False)
           .str.extract(r'(\d{4}/\d{2})')[0].fillna('نامشخص'))


# ################################################################################
# ################################################################################
# ################################################################################
def top_by(sub, agg):
    s = sub.groupby(DIS)[COST].agg(agg)
    if s.empty:
        return '', 0
    mx = s.max()
    return '، '.join(map(str, s[s == mx].index)), mx

def summarize(data, key, extra_key=None):
    # Aggregate disease groups once; sorted group keys preserve the original
    # groupby order and first-maximum tie breaking.
    grouped = data.groupby([key, DIS], sort=True)[COST].agg(['count', 'sum']).reset_index()
    totals = grouped.groupby(key, sort=True)['sum'].sum()
    counts = grouped.groupby(key, sort=True)['count'].sum()
    rows = []
    for k, sub in grouped.groupby(key, sort=True):
        dom_count = sub['count'].max()
        exp_cost = sub['sum'].max()
        dom = '، '.join(map(str, sub.loc[sub['count'].eq(dom_count), DIS]))
        exp = '، '.join(map(str, sub.loc[sub['sum'].eq(exp_cost), DIS]))
        total = totals[k]
        row = {key: k,
               'تعداد بیماری ثبت شده': counts[k],
               'مجموع هزینه‌های پرداخت شده (ریال)': total,
               'پرتکرارترین بیماری‌ها': dom,
               'پرهزینه‌ترین بیماری‌ها': exp,
               'سهم درصد پرهزینه‌ترین بیماری‌ها': round(exp_cost / total, 4) if total else 0}
        if extra_key:
            row['تعداد افراد تحت پوشش'] = data.loc[data[key].eq(k), extra_key].nunique()
        rows.append(row)
    return pd.DataFrame(rows)

def ranked(data, by= 'مجموع هزینه‌های پرداخت شده (ریال)'):
    """مرتب‌سازی نزولی + افزودن ستون رتبه به‌عنوان اولین ستون."""
    out = data.sort_values(by, ascending=False).reset_index(drop=True)
    out.insert(0, 'رتبه', out[by].rank(method='min', ascending=False).astype(int))
    return out

def monthly(key, values=None):
    """جدول ماهانه: سطرها شناسه، ستون‌ها ماه؛ تعداد خدمات یا مجموع هزینه."""
    if values is None:
        out = pd.pivot_table(df, index=key, columns=MON, values=DIS,
                             aggfunc='count', fill_value=0)
        total = 'جمع کل تعداد'
    else:
        out = pd.pivot_table(df, index=key, columns=MON, values=values,
                             aggfunc='sum', fill_value=0)
        total = 'جمع کل هزینه (ریال)'
    out = out.reindex(sorted(out.columns), axis=1)
    out[total] = out.sum(axis=1)
    return out.rename_axis(columns=None).reset_index()

# ################################################################################
# ################################################################################
# ################################################################################
sum_org = summarize(df, ID_ORG, ID_PER)
sum_per = summarize(df, ID_PER)

distinct_org = (df.groupby(ID_ORG)[DIS]
                  .nunique()
                  .reset_index(name='تعداد بیماری‌های متمایز'))

distinct_per = (df.groupby(ID_PER)[DIS]
                  .nunique()
                  .reset_index(name='تعداد بیماری‌های متمایز'))

dis_count = (df[DIS].value_counts()
             .rename_axis(DIS)
             .reset_index(name='تعداد ثبت‌شده'))

dis_cost = (df.groupby(DIS)[COST].sum()
            .reset_index()
            .rename(columns={COST: 'هزینه مجموع'}))

# #############################
mon_dis = pd.crosstab(df[MON], df[DIS])
mon_dis = mon_dis.reindex(sorted(mon_dis.index))
mon_dis['جمع کل تعداد'] = mon_dis.sum(axis=1)
mon_dis['مجموع هزینه پرداخت شده (ریال)'] = df.groupby(MON)[COST].sum()
mon_dis = mon_dis.rename_axis(columns=None).reset_index()
# #############################

sheets = {
    "خلاصه ماهانه": mon_dis,
    "خلاصه پرسنلی":        sum_org,
    "خلاصه کدملی":         sum_per,
    "تعداد بیماری متمایز-پرسنلی": distinct_org,
    "تعداد بیماری متمایز-کدملی": distinct_per,
    "مرتب‌شده پرسنلی-مجموع هزینه":     ranked(sum_org),
    "مرتب‌شده کدملی-مجموع هزینه":      ranked(sum_per),
    "مرتب‌شده پرسنلی-تعداد بیماری":     ranked(sum_org,     'تعداد بیماری ثبت شده'),
    "مرتب‌شده کدملی-تعداد بیماری":      ranked(sum_per,     'تعداد بیماری ثبت شده'),
    "تعداد ماهانه بیماری-پرسنلی":  monthly(ID_ORG),
    "تعداد ماهانه بیماری-کدملی":   monthly(ID_PER),
    "مجموع هزینه ماهانه-پرسنلی":        monthly(ID_ORG, COST),
    "مجموع هزینه ماهانه-کدملی":         monthly(ID_PER, COST),
    "فراوانی بیماری-پرسنلی": pd.crosstab(df[ID_ORG], df[DIS]).reset_index(),
    "فراوانی بیماری-کدملی":  pd.crosstab(df[ID_PER], df[DIS]).reset_index(),
    "هزینه تجمیعی بیماری-پرسنلی":   pd.pivot_table(df, index=ID_ORG, columns=DIS, values=COST,
                                          aggfunc='sum', fill_value=0).reset_index(),
    "هزینه تجمیعی بیماری-کدملی":    pd.pivot_table(df, index=ID_PER, columns=DIS, values=COST,
                                          aggfunc='sum', fill_value=0).reset_index(),
    "مرتب‌شده بیماری‌ها-تعداد": ranked(dis_count, by='تعداد ثبت‌شده'),
    "مرتب‌شده بیماری‌ها-مجموع هزینه": ranked(dis_cost,  by='هزینه مجموع'),
}

# ---------------- استایل ----------------
FA_RE = re.compile(r'[\u0600-\u06FF\uFB50-\uFDFF\uFE70-\uFEFF]')
EN_RE = re.compile(r'[A-Za-z]')

# فرمت اعداد با ارقام فارسی (کد زبان فارسی: 2000429)
FMT_INT     = '[$-2000429]#,##0'
FMT_PERCENT = '[$-2000429]0.00%'

F_FA   = Font(name='B Nazanin', size=11)
F_FA_B = Font(name='B Nazanin', size=11, bold=True)
F_EN   = Font(name='Cambria',   size=10)
F_EN_B = Font(name='Cambria',   size=10, bold=True)

def pick_font(v, bold=False):
    if isinstance(v, str) and EN_RE.search(v) and not FA_RE.search(v):
        return F_EN_B if bold else F_EN
    return F_FA_B if bold else F_FA

def style_sheet(ws, headers):
    # ۱) بازنویسی استایل پایه تا فونت پیش‌فرض کارپوشه هم عوض شود
    align_head = Alignment(horizontal='center', vertical='center', wrap_text=True, readingOrder=2)
    align_body = Alignment(horizontal='center', vertical='center', wrap_text=False, readingOrder=2)
    widths = [8] * len(headers)
    percent_cols = ['سهم' in str(col_name) for col_name in headers]
    for row in ws.iter_rows():
        for cell in row:
            head = (cell.row == 1)
            cell.font = pick_font(cell.value, bold=head)
            cell.alignment = align_head if head else align_body
            # ۲) فرمت عددی فارسی فقط برای سلول‌های عددی بدنه
            if not head and isinstance(cell.value, (int, float)):
                col = cell.column - 1
                cell.number_format = FMT_PERCENT if percent_cols[col] else FMT_INT
            if cell.value is not None:
                col = cell.column - 1
                widths[col] = max(widths[col], len(str(cell.value)))

    ws.sheet_view.rightToLeft = True
    ws.freeze_panes = 'A2'
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = min(w + 5, 40)

# ---------------- خروجی ----------------
with pd.ExcelWriter("export.xlsx", engine="openpyxl") as writer:
    for name, data in sheets.items():
        data.to_excel(writer, sheet_name=name, index=False)

    wb = writer.book
    # جایگزینی فونت استایل Normal (ریشهٔ مشکل «فونت اعمال نمی‌شود»)
    if 'Normal' in wb.named_styles:
        wb._named_styles[wb.style_names.index('Normal')].font = F_FA

    for name, data in sheets.items():
        style_sheet(wb[name], list(data.columns))

t2 = time()
print('Run Time =', round(t2 - t1, 2), 's')