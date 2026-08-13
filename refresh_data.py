"""
CBRE HK Office Dashboard — Data Refresh Script
Reads all 4 Excel files and outputs dashboard_data.json
Run this manually or via schedule to keep the dashboard current.
"""

import pandas as pd
import json
import math
import warnings
import os
from datetime import datetime

warnings.filterwarnings('ignore')

BASE = os.path.dirname(os.path.abspath(__file__))
# Auto-detect: if Excel files are in the same folder (GitHub Actions), use BASE
# Otherwise fall back to sibling "Office Databases" folder (local Cowork)
if os.path.exists(os.path.join(BASE, "Hong Kong Office Database 3.0.xlsm")):
    DB = BASE
else:
    DB = os.path.join(os.path.dirname(BASE), "Office Databases")
OUT  = os.path.join(BASE, "dashboard_data.json")

def safe(v):
    """Convert NaN/inf to None for JSON."""
    if v is None: return None
    try:
        if math.isnan(float(v)) or math.isinf(float(v)): return None
        return round(float(v), 4)
    except: return str(v) if pd.notna(v) else None

def read_xl(name, **kwargs):
    return pd.read_excel(os.path.join(DB, name), **kwargs)

print(f"[{datetime.now():%Y-%m-%d %H:%M}] Starting data refresh...")

data = {}

# ══════════════════════════════════════════════════════════════════════
# 1. RENT & VACANCY — Hong Kong Office Database 3.0.xlsx
# ══════════════════════════════════════════════════════════════════════
print("  Reading Office Database...")
db_path = os.path.join(DB, "Hong Kong Office Database 3.0.xlsm")

# ── Current snapshot (MthSum) ──
df_mth = pd.read_excel(db_path, sheet_name='MthSum', header=None)
snap_date = str(df_mth.iloc[0, 1]) if pd.notna(df_mth.iloc[0, 1]) else "Latest"
snap = df_mth.iloc[3:, [0,1,4,10,13,19,20,21,22]].copy()
snap.columns = ['Location','NER','Vacancy_pct',
                'MOM_RentalChg','QOQ_RentalChg','YTD_RentalChg','YTD_VacChg',
                'YOY_RentalChg','YOY_VacChg']
snap = snap[snap['Location'].notna()]
for c in snap.columns[1:]:
    snap[c] = pd.to_numeric(snap[c], errors='coerce')

key_districts = ['Central','Admiralty & Sheung Wan','Wan Chai & Causeway Bay',
                 'Hong Kong East','Wong Chuk Hang','Tsim Sha Tsui',
                 'Tsim Sha Tsui East','Tsim Sha Tsui West','Kowloon East','Kowloon Others','New Territories']
snap_key = snap[snap['Location'].isin(key_districts)].copy()
snap_key['vac'] = (snap_key['Vacancy_pct']*100).round(1)
snap_key['ytd'] = (snap_key['YTD_RentalChg']*100).round(1)
snap_vac = snap_key.sort_values('vac', ascending=True)
snap_ytd = snap_key.sort_values('ytd', ascending=True)

def _rent_chg_series(col):
    df = snap_key.copy()
    df['v'] = (df[col]*100).round(2)
    df = df.sort_values('v', ascending=True)
    return {'labels': df['Location'].tolist(), 'values': df['v'].tolist()}

rent_change_by_loc = {
    'ytd': _rent_chg_series('YTD_RentalChg'),
    'yoy': _rent_chg_series('YOY_RentalChg'),
    'qoq': _rent_chg_series('QOQ_RentalChg'),
    'mom': _rent_chg_series('MOM_RentalChg'),
}

overall_row = snap[snap['Location']=='Overall Hong Kong'].iloc[0]
hki_row     = snap[snap['Location']=='Hong Kong Island'].iloc[0]
kln_row     = snap[snap['Location']=='Kowloon'].iloc[0]
nt_row      = snap[snap['Location']=='New Territories'].iloc[0]

# ── Output_Rent: NER per location ─────────────────────────────────────
# Row 9 = location names, Row 10 = metric names, data from row 11.
# Year in col 0, Month (text) in col 1.
df_rent      = pd.read_excel(db_path, sheet_name='Output_Rent', header=None)
_rent_loc    = df_rent.iloc[9]
_rent_metric = df_rent.iloc[10]
_NER_METRIC  = 'Net Effective Rent (Market) (HK$psf)'
_ner_col_idx = {str(_rent_loc[c]).strip(): c
                for c in range(2, df_rent.shape[1])
                if str(_rent_metric[c]).strip() == _NER_METRIC
                and str(_rent_loc[c]).strip() not in ('nan', '', 'Weighted Average Rent')}

# ── Output_SAV: Net Absorption per location ────────────────────────────
# Row 7 = location names, Row 8 = metric names, data from row 9.
df_sav2     = pd.read_excel(db_path, sheet_name='Output_SAV', header=None)
_sav_loc    = df_sav2.iloc[7]
_sav_metric = df_sav2.iloc[8]
_NA_METRIC  = 'Net Absorption (Monthly)'
_na_col_idx = {str(_sav_loc[c]).strip(): c
               for c in range(2, df_sav2.shape[1])
               if str(_sav_metric[c]).strip() == _NA_METRIC
               and str(_sav_loc[c]).strip() not in ('nan', '')}

_month_num = {'January':1,'February':2,'March':3,'April':4,'May':5,'June':6,
              'July':7,'August':8,'September':9,'October':10,'November':11,'December':12}

def _build_out_frame(df, col_idx, data_start_row):
    cols  = [0, 1] + list(col_idx.values())
    names = ['Year', 'Month'] + list(col_idx.keys())
    frame = df.iloc[data_start_row:, cols].copy()
    frame.columns = names
    frame['Year']  = pd.to_numeric(frame['Year'], errors='coerce')
    frame['Month'] = frame['Month'].map(
        lambda m: _month_num.get(str(m).strip(), None) if isinstance(m, str)
                  else (int(m) if pd.notna(m) else None))
    frame = frame.dropna(subset=['Year', 'Month'])
    for c in frame.columns[2:]: frame[c] = pd.to_numeric(frame[c], errors='coerce')
    frame['Year']    = frame['Year'].astype(int)
    frame['Month']   = frame['Month'].astype(int)
    frame['Quarter'] = ((frame['Month'] - 1) // 3 + 1)
    frame['Date']    = pd.to_datetime(
        frame['Year'].astype(str) + '-' + frame['Month'].astype(str).str.zfill(2) + '-01')
    frame = frame.drop_duplicates(subset=['Year','Month'], keep='first')
    return frame.sort_values(['Year','Month'])

nr     = _build_out_frame(df_rent, _ner_col_idx, 11)  # NER  from Output_Rent
na_out = _build_out_frame(df_sav2, _na_col_idx,  9)   # NA   from Output_SAV

# ── Patch rent_change_by_loc: QOQ + missing YOY from Output sheet ─────
# QOQ = (current_month_NER - NER_3months_ago) / NER_3months_ago
# YOY = (current_month_NER - NER_12months_ago) / NER_12months_ago
# Use the latest row in nr as "current", back-index 3 and 12 rows for comparison.
_nr_sorted = nr.sort_values(['Year','Month']).reset_index(drop=True)
if len(_nr_sorted) >= 13:
    _cur  = _nr_sorted.iloc[-1]
    _m12  = _nr_sorted.iloc[-13]  # 12 months prior

    def _pct_chg(cur_val, prev_val):
        try:
            c, p = float(cur_val), float(prev_val)
            if p and not pd.isna(c) and not pd.isna(p):
                return round((c - p) / p * 100, 2)
        except (TypeError, ValueError):
            pass
        return None

    _snap_month = int(_cur['Month'])
    _is_qtr_end = _snap_month in (3, 6, 9, 12)

    # QOQ: only compute when snapshot is a quarter-end month
    if _is_qtr_end and len(_nr_sorted) >= 4:
        _m3 = _nr_sorted.iloc[-4]   # 3 months prior
        _qoq_map = {loc: _pct_chg(_cur[loc], _m3[loc])
                    for loc in key_districts if loc in _cur.index}
        _rows_q = [(loc, v) for loc, v in _qoq_map.items() if v is not None]
        _rows_q.sort(key=lambda x: x[1])
        rent_change_by_loc['qoq'] = {'labels': [r[0] for r in _rows_q],
                                     'values': [r[1] for r in _rows_q]}
    else:
        rent_change_by_loc['qoq'] = None   # mid-quarter — disable button in dashboard

    # YOY: fill only entries missing from MthSum (e.g. TST West)
    _yoy_map = {loc: _pct_chg(_cur[loc], _m12[loc])
                for loc in key_districts if loc in _cur.index}
    yoy_d = rent_change_by_loc['yoy']
    existing_yoy = dict(zip(yoy_d['labels'], yoy_d['values']))
    for loc, v in _yoy_map.items():
        if loc not in existing_yoy and v is not None:
            existing_yoy[loc] = v
    rows_yoy = sorted(existing_yoy.items(), key=lambda x: x[1])
    rent_change_by_loc['yoy'] = {'labels': [r[0] for r in rows_yoy],
                                 'values': [r[1] for r in rows_yoy]}

annual_ner = (nr[nr['Year'] >= 2008]
              .sort_values('Month', ascending=False)
              .groupby('Year').first()
              .reset_index())

def avg_cols(df, cols):
    return [safe(row[cols].dropna().mean()) for _, row in df.iterrows()]

# Annual NER — use Output sheet area columns directly
hki_ner = [safe(v) for v in annual_ner['Hong Kong Island']]
kln_ner = [safe(v) for v in annual_ner['Kowloon']]
nt_ner  = [safe(v) for v in annual_ner['New Territories']]
ner_yrs = annual_ner['Year'].tolist()

# ── Per-submarket NER time series (for trend filter) ──────────────────
# Column names in Output sheet match submarket display names exactly.
SUBMARKET_COLS = {sm: sm for sm in [
    'Central', 'Admiralty & Sheung Wan', 'Wan Chai & Causeway Bay',
    'Hong Kong East', 'Wong Chuk Hang',
    'Tsim Sha Tsui', 'Tsim Sha Tsui East', 'Tsim Sha Tsui West',
    'Kowloon East', 'Kowloon Others', 'New Territories',
]}
ner_by_submarket = {}
for label in SUBMARKET_COLS:
    if label in annual_ner.columns:
        ner_by_submarket[label] = [safe(v) for v in annual_ner[label]]

# NOTE: Per-submarket vacancy time series are built AFTER the SAV series and
# MthSum snapshot are read (further below). See "vac_by_submarket" block.

# ── Overall vacancy % (SAV) ──
df_sav = pd.read_excel(db_path, sheet_name='SAV', header=None)
sav = df_sav.iloc[7:, [0,1,12]].copy()
sav.columns = ['Year','Month','VacPct']
sav['Year'] = pd.to_numeric(sav['Year'], errors='coerce')
sav['VacPct'] = pd.to_numeric(sav['VacPct'], errors='coerce')
sav = sav.dropna(subset=['Year','Month','VacPct'])
month_ord = {m:i for i,m in enumerate(['January','February','March','April','May','June',
                                        'July','August','September','October','November','December'],1)}
sav['MN'] = sav['Month'].map(month_ord)
sav_ann = sav[sav['Year']>=2008].groupby('Year').apply(lambda x: x.nlargest(1,'MN')).reset_index(drop=True)

hki_r, kln_r, nt_r = (float(hki_row['Vacancy_pct']), float(kln_row['Vacancy_pct']), float(nt_row['Vacancy_pct']))
all_r = float(overall_row['Vacancy_pct'])
hki_vac = [safe(v*hki_r/all_r*100) for v in sav_ann['VacPct']]
kln_vac = [safe(v*kln_r/all_r*100) for v in sav_ann['VacPct']]
nt_vac  = [safe(v*nt_r/all_r*100)  for v in sav_ann['VacPct']]
vac_yrs = sav_ann['Year'].tolist()

# ── Per-submarket vacancy time series ────────────────────────────────────
# Method: scale the pre-calculated SAV overall-HK trend by each submarket's
# current vacancy ratio from MthSum — same approach used for HKI/KLN/NT above.
# All source values come directly from pre-calculated Excel sheets (no custom
# aggregation from raw data).
vac_by_submarket = {}
for sm in SUBMARKET_COLS.keys():
    sm_row = snap[snap['Location'] == sm]
    if len(sm_row) == 0:
        continue
    sm_r = float(sm_row.iloc[0]['Vacancy_pct'])
    vac_by_submarket[sm] = [safe(v * sm_r / all_r * 100) for v in sav_ann['VacPct']]
vac_yrs_sm = vac_yrs   # same year axis as overall

# ── Building-level snapshot built after Space + Rental are loaded below ──
# (Output2 sheet was removed in DB 3.0; bld/top_vac/top_ner defined further down)

# ── Bldg list — source of truth for District / Submarket / Area ──────
df_bl = pd.read_excel(db_path, sheet_name='Bldg List', header=None)
df_bl.columns = ['RefNo','BldgName','BldgCHI','District','Submarket','Area','Grade','Ownership',
                 'Landlord','Portfolio','PortGroup','GFA','Efficiency','NFA','FlrPlateGFA','FlrPlateNFA',
                 'GrossNet','Location','Strata','CompYear','CompDate','Age','AgeBand','NewSupply','Status',
                 'EffOn','EffOn2','ExpOn','Lat','Long','Polygon']
bldg_info = df_bl[df_bl['BldgName'].notna() & (df_bl['BldgName']!='Building Name (ENG)')].copy()
# Keep only Valid buildings (col Y = Status) — used for current snapshots,
# portfolio groups, building map, and the main geography reference.
bldg_info = bldg_info[bldg_info['Status'] == 'Valid'].copy()
# Sort so rows WITH a PortGroup value come first within each RefNo group,
# ensuring drop_duplicates keeps the row that carries the portfolio assignment.
bldg_info = bldg_info.sort_values('PortGroup', na_position='last')
bldg_info = bldg_info.drop_duplicates('RefNo', keep='first')

# ── All-status geography lookup (for historical Space-based series) ────────
# Expired/Demolished buildings still have historical Space records and
# contribute to SAV's Overall HK NA.  Excluding them from the submarket map
# causes the submarket sum to deviate from Overall HK.  This wider lookup
# retains their Submarket/District so df_bld_mth can include them correctly.
bldg_info_all = df_bl[df_bl['BldgName'].notna() & (df_bl['BldgName']!='Building Name (ENG)')].copy()
bldg_info_all = bldg_info_all.sort_values('PortGroup', na_position='last')
bldg_info_all = bldg_info_all.drop_duplicates('RefNo', keep='first')

# ── Rental sheet — per-building monthly Net Effective Rent ──────────────
# Used to build per-portfolio NER time series from actual data
# (replaces proportional-scaling approximation used previously).
df_rental = pd.read_excel(db_path, sheet_name='Rental', header=None)
df_rental.columns = ['Date', 'RefNo', 'RentalType', 'Value']
df_rental = df_rental.iloc[1:]
df_rental['Date']  = pd.to_datetime(df_rental['Date'],  errors='coerce')
df_rental['Value'] = pd.to_numeric(df_rental['Value'],  errors='coerce')
df_rental_ner = df_rental[df_rental['RentalType'] == 'Net Effective Rent'].copy()
df_rental_ner = df_rental_ner.dropna(subset=['Date', 'Value'])
df_rental_ner['Year']    = df_rental_ner['Date'].dt.year
df_rental_ner['Month']   = df_rental_ner['Date'].dt.month
df_rental_ner['Quarter'] = df_rental_ner['Date'].dt.quarter

# ── Merge Rental_Revision (pre-basket NER back-fills) ────────────────────
# Buildings enter stock before they join the rental basket.
# Rental_Revision contains back-calculated NER for those pre-basket months
# (e.g. if a building completes Jan 2020 but joins basket Jul 2020, the
#  Jan–Jun 2020 NER is derived from district m-o-m change and stored here).
# Concat: basket data first (priority), RR second; drop_duplicates keeps
# the basket row wherever both sources have the same (Date, RefNo).
try:
    df_rr = pd.read_excel(db_path, sheet_name='Rental_Revision', header=0, usecols=[0, 1, 2])
    df_rr.columns = ['Date', 'RefNo', 'Value']
    df_rr['Date']    = pd.to_datetime(df_rr['Date'],  errors='coerce')
    df_rr['Value']   = pd.to_numeric(df_rr['Value'],  errors='coerce')
    df_rr = df_rr.dropna(subset=['Date', 'Value'])
    df_rr['Year']    = df_rr['Date'].dt.year
    df_rr['Month']   = df_rr['Date'].dt.month
    df_rr['Quarter'] = df_rr['Date'].dt.quarter
    df_rental_ner = (pd.concat([df_rental_ner, df_rr], ignore_index=True)
                       .drop_duplicates(subset=['Date', 'RefNo'], keep='first'))
    print(f"    Rental_Revision: {len(df_rr)} records merged ({df_rr['RefNo'].nunique()} buildings)")
except Exception as e:
    print(f"    Rental_Revision not found or error: {e}")

# ── Latest NER per building (used for building-level snapshot) ────────────
_bld_ner = (df_rental_ner.sort_values('Date')
             .drop_duplicates('RefNo', keep='last')
             [['RefNo','Value']].rename(columns={'Value':'NER'}))

# ── Vacant Space (sq.ft. NFA) time series — from Space sheet ─────────────
# Source: Space sheet SpaceType='Vacant', joined to Bldg List for geography.
# Method: year-end snapshot per building (last record each year), then sum
# by Area / Submarket / PortGroup.  Overall HK comes directly from SAV col 4.
print("  Reading Space sheet (this may take a moment)...")
df_space = pd.read_excel(db_path, sheet_name='Space', header=None)
df_space.columns = ['Date', 'RefNo', 'SpaceType', 'Value']
df_space = df_space.iloc[1:]   # drop header row
df_space['Date']  = pd.to_datetime(df_space['Date'],  errors='coerce')
df_space['Value'] = pd.to_numeric(df_space['Value'], errors='coerce')
df_space_nfa = df_space[df_space['SpaceType'] == 'Net Floor Area'].dropna(subset=['Date', 'Value']).copy()

# Extract Future Vacancy and Shadow Space before filtering to Vacant-only
def _latest_per_bldg(stype, col_name):
    return (df_space[df_space['SpaceType'] == stype]
            .dropna(subset=['Date', 'Value'])
            [['Date','RefNo','Value']]
            .sort_values('Date')
            .drop_duplicates('RefNo', keep='last')
            .rename(columns={'Value': col_name})[['RefNo', col_name]])

_bld_futvac = _latest_per_bldg('Future', 'FutureVac')
_bld_shadow = _latest_per_bldg('Shadow', 'ShadowVac')

df_space = df_space[df_space['SpaceType'] == 'Vacant'].dropna(subset=['Date', 'Value'])

# ── Building-level snapshot: NER (Rental) + Vacant / FutureVac / Shadow (Space) + NFA (Bldg List) ──
_bld_vac = (df_space[['Date','RefNo','Value']]
             .sort_values('Date')
             .drop_duplicates('RefNo', keep='last')
             .rename(columns={'Value':'Vacant'})[['RefNo','Vacant']])
_bld_base = bldg_info[['RefNo','BldgName','NFA']].copy()
_bld_base['NFA'] = pd.to_numeric(_bld_base['NFA'], errors='coerce')
bld = (_bld_base
       .merge(_bld_ner,    on='RefNo', how='left')
       .merge(_bld_vac,    on='RefNo', how='left')
       .merge(_bld_futvac, on='RefNo', how='left')
       .merge(_bld_shadow, on='RefNo', how='left')
       .rename(columns={'BldgName':'Name'}))
bld['VacPct']    = bld['Vacant'] / bld['NFA'].replace(0, float('nan'))
bld['_TotalVac'] = bld[['Vacant','FutureVac','ShadowVac']].fillna(0).sum(axis=1)
top_vac = bld[bld['_TotalVac']>0].nlargest(15,'_TotalVac')
top_ner = bld[bld['NER']>0].nlargest(15,'NER')

# Join geography from Bldg List (Valid buildings only — for vacancy/NFA snapshots)
sp_bl = bldg_info[['RefNo', 'Area', 'Submarket', 'District', 'PortGroup']].copy()
sp_bl['PortGroup_s'] = sp_bl['PortGroup'].astype(str).replace({'nan': ''})
sp_bl['District_s']  = sp_bl['District'].astype(str).replace({'nan': ''})
df_space = df_space.merge(sp_bl, on='RefNo', how='left')

# ── All-status geography for NA (Expired/Demolished buildings included) ───
# Use bldg_info_all so that buildings which have been expired/demolished still
# retain their Submarket/District label when computing per-submarket NA.
# Without this the submarket sum cannot reconcile to SAV's Overall HK figure.
sp_bl_all = bldg_info_all[['RefNo', 'Submarket', 'District']].copy()
sp_bl_all['District_s'] = sp_bl_all['District'].astype(str).replace({'nan': ''})

# ── Per-building monthly Net Absorption basis ────────────────────────────
# Same source as Power Pivot [Net Absorption (Monthly)] measure:
#   NA = ΔOccupied = ΔStock − ΔVacant  →  NA = −ΔVacant + max(0, ΔNFA)
# Build _vac_b from the raw Space Vacant records (before Valid-only merge)
# so Expired/Demolished buildings are included.
_vac_raw = df_space[['Date','RefNo','Value']].rename(columns={'Value':'Vacant'})
_vac_b = _vac_raw.merge(sp_bl_all, on='RefNo', how='left')
_nfa_b = df_space_nfa[['Date','RefNo','Value']].rename(columns={'Value':'NFA'})
df_bld_mth = _vac_b.merge(_nfa_b, on=['Date','RefNo'], how='inner')
df_bld_mth['Submarket_s'] = df_bld_mth['Submarket'].astype(str).replace({'nan':''})
df_bld_mth = df_bld_mth.sort_values(['RefNo','Date'])
df_bld_mth['dVacant'] = df_bld_mth.groupby('RefNo')['Vacant'].diff()
df_bld_mth['dNFA']    = df_bld_mth.groupby('RefNo')['NFA'].diff()
df_bld_mth['NA']      = -df_bld_mth['dVacant'] + df_bld_mth['dNFA'].clip(lower=0)
df_bld_mth = df_bld_mth.dropna(subset=['dVacant'])   # drop first row per building
df_bld_mth['Year']    = df_bld_mth['Date'].dt.year
df_bld_mth['Month']   = df_bld_mth['Date'].dt.month
df_bld_mth['Quarter'] = df_bld_mth['Date'].dt.quarter
df_bld_mth['MLabel']  = df_bld_mth['Year'].astype(str) + '-' + df_bld_mth['Month'].astype(str).str.zfill(2)
df_bld_mth['QLabel']  = df_bld_mth['Year'].astype(str) + ' Q' + df_bld_mth['Quarter'].astype(str)

# Year-end snapshot per building: keep the latest record in each calendar year
df_space['Year']  = df_space['Date'].dt.year
df_space['Month'] = df_space['Date'].dt.month
sp_yr = df_space[df_space['Year'] >= 2008].copy()
sp_last = (sp_yr.groupby(['RefNo', 'Year'])
               .apply(lambda x: x.nlargest(1, 'Month'))
               .reset_index(drop=True))

# Overall HK vacant sf — use SAV sheet pre-calculated value (col 4)
sav_vacsf = df_sav.iloc[7:, [0, 1, 4]].copy()
sav_vacsf.columns = ['Year', 'Month', 'VacSF']
sav_vacsf['Year']  = pd.to_numeric(sav_vacsf['Year'],  errors='coerce')
sav_vacsf['VacSF'] = pd.to_numeric(sav_vacsf['VacSF'], errors='coerce')
sav_vacsf = sav_vacsf.dropna(subset=['Year', 'Month', 'VacSF'])
sav_vacsf['MN'] = sav_vacsf['Month'].map(month_ord)
sav_vacsf_ann = (sav_vacsf[sav_vacsf['Year'] >= 2008]
                 .groupby('Year')
                 .apply(lambda x: x.nlargest(1, 'MN'))
                 .reset_index(drop=True))
vac_sf_overall = [safe(v) for v in sav_vacsf_ann['VacSF']]   # aligned with vac_yrs

# Area-level (HKI / Kowloon / NT)
area_vac = sp_last.groupby(['Area', 'Year'])['Value'].sum().reset_index()
area_years = sorted(sp_last['Year'].unique().tolist())
vac_sf_by_area = {}
for area in ['Hong Kong Island', 'Kowloon', 'New Territories']:
    subset = area_vac[area_vac['Area'] == area].set_index('Year')
    vac_sf_by_area[area] = [safe(subset.loc[y, 'Value']) if y in subset.index else None
                            for y in area_years]

# Per-submarket
sm_vac = sp_last.groupby(['Submarket', 'Year'])['Value'].sum().reset_index()
vac_sf_by_submarket = {}
for sm in sp_last['Submarket'].dropna().unique():
    subset = sm_vac[sm_vac['Submarket'] == sm].set_index('Year')
    vac_sf_by_submarket[str(sm)] = [safe(subset.loc[y, 'Value']) if y in subset.index else None
                                    for y in area_years]

# Per-portfolio group
pg_vac = (sp_last[sp_last['PortGroup_s'].str.len() > 0]
          .groupby(['PortGroup_s', 'Year'])['Value'].sum().reset_index())
vac_sf_by_portfolio = {}
for pg_name in sp_last['PortGroup_s'].dropna().unique():
    if not pg_name: continue
    subset = pg_vac[pg_vac['PortGroup_s'] == pg_name].set_index('Year')
    vac_sf_by_portfolio[str(pg_name)] = [safe(subset.loc[y, 'Value']) if y in subset.index else None
                                         for y in area_years]

# Per-district annual vacant SF
dist_vac = sp_last.groupby(['District_s', 'Year'])['Value'].sum().reset_index()
vac_sf_by_district = {}
for d_name in sp_last['District_s'].dropna().unique():
    if not d_name: continue
    sub = dist_vac[dist_vac['District_s'] == d_name].set_index('Year')
    vac_sf_by_district[str(d_name)] = [safe(sub.loc[y, 'Value']) if y in sub.index else None
                                        for y in area_years]

print(f"    Vacant SF series: {len(vac_sf_by_submarket)} submarkets, "
      f"{len(vac_sf_by_portfolio)} portfolio groups, {len(vac_sf_by_district)} districts, "
      f"years {area_years[0]}\u2013{area_years[-1]}")

# Build lookup: RefNo → {district, submarket, area, grade}
bldg_map = {}
for _, r in bldg_info.iterrows():
    if pd.notna(r['RefNo']):
        bldg_map[str(r['RefNo'])] = {
            'district':  str(r['District'])  if pd.notna(r['District'])  else '',
            'submarket': str(r['Submarket']) if pd.notna(r['Submarket']) else '',
            'area':      str(r['Area'])      if pd.notna(r['Area'])      else '',
            'grade':     str(r['Grade'])     if pd.notna(r['Grade'])     else '',
        }
bldg_info['Lat'] = pd.to_numeric(bldg_info['Lat'], errors='coerce')
bldg_info['Long'] = pd.to_numeric(bldg_info['Long'], errors='coerce')
bldg_info['NFA'] = pd.to_numeric(bldg_info['NFA'], errors='coerce')
bldg_info['CompYear'] = pd.to_numeric(bldg_info['CompYear'], errors='coerce')
# Merge NER from Output2
ner_map = bld[['RefNo','NER','Vacant','FutureVac','ShadowVac','VacPct']].copy()
bldg_merged = bldg_info.merge(ner_map, on='RefNo', how='left')

# ── Portfolio group current NER & vacancy ─────────────────────────────
bldg_merged['NER_n']    = pd.to_numeric(bldg_merged['NER'],    errors='coerce')
bldg_merged['VacPct_n'] = pd.to_numeric(bldg_merged['VacPct'], errors='coerce')
bldg_merged['PortGroup_s'] = bldg_merged['PortGroup'].astype(str).replace({'nan':''})
port_grp = (bldg_merged[bldg_merged['PortGroup_s'].str.len()>0]
            .groupby('PortGroup_s', sort=True)
            .agg(ner=('NER_n','mean'), vacpct=('VacPct_n','mean'))
            .reset_index())
portfolio_groups = [
    {'name': str(r['PortGroup_s']),
     'ner':    safe(r['ner']),
     'vacpct': round(float(r['vacpct'])*100, 1) if pd.notna(r['vacpct']) else None}
    for _, r in port_grp.iterrows()
]

# ── Per-portfolio NER & vacancy time series ───────────────────────────────
# Same proportional-scaling method as submarkets: use overall pre-calculated
# trend as base, scaled by each portfolio's current ratio vs overall.
overall_ner_series = [safe(v) for v in annual_ner['Overall Hong Kong']]
current_overall_ner = float(overall_row['NER'])    # from MthSum
current_overall_vac = all_r                         # from MthSum (fraction)

# ── Per-portfolio NER from Rental sheet (actual data) ─────────────────
_pg_map = dict(zip(bldg_info['RefNo'].astype(str), bldg_info['PortGroup'].astype(str)))
df_rental_ner['PortGroup'] = df_rental_ner['RefNo'].astype(str).map(_pg_map)
df_rental_ner_pg = df_rental_ner[
    df_rental_ner['PortGroup'].notna() & (df_rental_ner['PortGroup'] != 'nan')
].copy()

# Annual: last-month snapshot per building per year, then mean across portfolio
_rner_ann = (df_rental_ner_pg[df_rental_ner_pg['Year'] >= 2008]
             .groupby(['PortGroup', 'Year', 'RefNo'])
             .apply(lambda x: x.nlargest(1, 'Month'))
             .reset_index(drop=True)
             .groupby(['PortGroup', 'Year'])['Value'].mean().reset_index())

ner_by_portfolio = {}
for pg in portfolio_groups:
    nm = pg['name']
    sub = _rner_ann[_rner_ann['PortGroup'] == nm].set_index('Year')
    ner_by_portfolio[nm] = [safe(sub.loc[y, 'Value']) if y in sub.index else None
                            for y in ner_yrs]

# ── Current NER & vacancy snapshot per Bldg List submarket ────────────
bldg_merged['NFA_n'] = pd.to_numeric(bldg_merged['NFA'], errors='coerce')
sub_snap = (bldg_merged[bldg_merged['Submarket'].notna()]
            .groupby('Submarket', sort=True)
            .agg(ner=('NER_n','mean'), vacpct=('VacPct_n','mean'),
                 total_nfa=('NFA_n','sum'))
            .reset_index())
submarket_snapshot = [
    {'name':      str(r['Submarket']),
     'ner':       safe(r['ner']),
     'vacpct':    round(float(r['vacpct'])*100, 1) if pd.notna(r['vacpct']) else None,
     'total_nfa': safe(r['total_nfa'])}
    for _, r in sub_snap.iterrows()
    if pd.notna(r['ner'])
]

# ── Total NFA per portfolio group (for vacant-sf time series) ─────────
port_nfa = (bldg_merged[bldg_merged['PortGroup_s'].str.len()>0]
            .groupby('PortGroup_s')['NFA_n'].sum().reset_index())
port_nfa_map = dict(zip(port_nfa['PortGroup_s'], port_nfa['NFA_n']))
# Append total_nfa to portfolio_groups
for pg in portfolio_groups:
    pg['total_nfa'] = safe(port_nfa_map.get(pg['name'], 0))

# ── Per-portfolio Vacancy % from Space + NFA (actual data) ─────────────
vac_by_portfolio = {}
for pg in portfolio_groups:
    nm = pg['name']
    total_nfa = port_nfa_map.get(nm, 0)
    if not total_nfa:
        continue
    pg_sub = pg_vac[pg_vac['PortGroup_s'] == nm].set_index('Year')
    vac_by_portfolio[nm] = [
        safe(pg_sub.loc[y, 'Value'] / total_nfa * 100) if y in pg_sub.index else None
        for y in area_years
    ]

# ── District-level current snapshot ─────────────────────────────────────
dist_snap = (bldg_merged[bldg_merged['District'].notna()]
             .groupby('District', sort=True)
             .agg(ner=('NER_n','mean'), vacpct=('VacPct_n','mean'), total_nfa=('NFA_n','sum'))
             .reset_index())
district_snapshot = [
    {'name': str(r['District']),
     'ner':  safe(r['ner']),
     'vacpct': round(float(r['vacpct'])*100, 1) if pd.notna(r['vacpct']) else None,
     'total_nfa': safe(r['total_nfa'])}
    for _, r in dist_snap.iterrows()
]

# ── District NFA map (for vacancy % calculation) ──────────────────────
dist_nfa = (bldg_merged[bldg_merged['District'].notna()]
            .groupby('District')['NFA_n'].sum().reset_index())
dist_nfa_map = dict(zip(dist_nfa['District'], dist_nfa['NFA_n']))

# ── District NER from Rental sheet (annual) ──────────────────────────
_dist_map = dict(zip(bldg_info['RefNo'].astype(str), bldg_info['District'].astype(str)))
df_rental_ner['District'] = df_rental_ner['RefNo'].astype(str).map(_dist_map)
df_rental_ner_dist = df_rental_ner[
    df_rental_ner['District'].notna() & (df_rental_ner['District'] != 'nan')
].copy()
_rner_dist_ann = (df_rental_ner_dist[df_rental_ner_dist['Year'] >= 2008]
                  .groupby(['District', 'Year', 'RefNo'])
                  .apply(lambda x: x.nlargest(1, 'Month'))
                  .reset_index(drop=True)
                  .groupby(['District', 'Year'])['Value'].mean().reset_index())
ner_by_district = {}
for d in district_snapshot:
    nm = d['name']
    sub = _rner_dist_ann[_rner_dist_ann['District'] == nm].set_index('Year')
    ner_by_district[nm] = [safe(sub.loc[y, 'Value']) if y in sub.index else None for y in ner_yrs]

# ── District Vacancy % from Space + NFA (annual) ─────────────────────
vac_by_district = {}
for d in district_snapshot:
    nm = d['name']
    total_nfa = dist_nfa_map.get(nm, 0)
    if not total_nfa: continue
    sub = dist_vac[dist_vac['District_s'] == nm].set_index('Year')
    vac_by_district[nm] = [
        safe(sub.loc[y, 'Value'] / total_nfa * 100) if y in sub.index else None
        for y in area_years
    ]

# ══════════════════════════════════════════════════════════════════════
# QUARTERLY & MONTHLY TIME SERIES (for granularity toggle)
# Quarterly: labels 'YYYY QN', 2008-present
# Monthly:   labels 'YYYY-MM', 2018-present
# ══════════════════════════════════════════════════════════════════════
print("  Building quarterly & monthly series...")

# ── NER quarterly / monthly (Output sheet — all locations, incl. September) ─
# NER is a rate/stock → quarterly = last month of quarter; monthly = direct.
# September gaps no longer exist — database now provides all 12 months.

nr_filt = nr[nr['Year'] >= 2008].copy()

nr_q = (nr_filt.groupby(['Year','Quarter'], sort=False)
               .apply(lambda x: x.nlargest(1,'Month'))
               .reset_index(drop=True)
               .sort_values(['Year','Quarter']))
nr_q['QLabel'] = nr_q['Year'].astype(str) + ' Q' + nr_q['Quarter'].astype(str)
ner_q_labels   = nr_q['QLabel'].tolist()
hki_ner_q      = [safe(v) for v in nr_q['Hong Kong Island']]
kln_ner_q      = [safe(v) for v in nr_q['Kowloon']]
nt_ner_q       = [safe(v) for v in nr_q['New Territories']]
overall_ner_q  = [safe(v) for v in nr_q['Overall Hong Kong']]
ner_by_sub_q   = {sm: [safe(v) for v in nr_q[sm]]
                  for sm in SUBMARKET_COLS if sm in nr_q.columns}

nr_m = nr_filt[nr_filt['Year'] >= 2018].sort_values(['Year','Month'])
nr_m['MLabel'] = nr_m['Date'].dt.strftime('%Y-%m')
ner_m_labels   = nr_m['MLabel'].tolist()
hki_ner_m      = [safe(v) for v in nr_m['Hong Kong Island']]
kln_ner_m      = [safe(v) for v in nr_m['Kowloon']]
nt_ner_m       = [safe(v) for v in nr_m['New Territories']]
overall_ner_m  = [safe(v) for v in nr_m['Overall Hong Kong']]
ner_by_sub_m   = {sm: [safe(v) for v in nr_m[sm]]
                  for sm in SUBMARKET_COLS if sm in nr_m.columns}
# ── Vacancy % quarterly / monthly (SAV sheet) ─────────────────────────
sav_filt = sav[sav['Year'] >= 2008].copy()
sav_filt['Quarter'] = ((sav_filt['MN'] - 1) // 3 + 1).astype(int)
sav_filt = sav_filt.sort_values(['Year','MN'])

sav_q = (sav_filt.groupby(['Year','Quarter'], sort=False)
                 .apply(lambda x: x.nlargest(1,'MN'))
                 .reset_index(drop=True)
                 .sort_values(['Year','Quarter']))
sav_q['QLabel'] = sav_q['Year'].astype(str) + ' Q' + sav_q['Quarter'].astype(str)
vac_q_labels    = sav_q['QLabel'].tolist()
hki_vac_q       = [safe(v*hki_r/all_r*100) for v in sav_q['VacPct']]
kln_vac_q       = [safe(v*kln_r/all_r*100) for v in sav_q['VacPct']]
nt_vac_q        = [safe(v*nt_r/all_r*100)  for v in sav_q['VacPct']]
vac_by_sub_q    = {}
for sm in SUBMARKET_COLS.keys():
    sm_row = snap[snap['Location'] == sm]
    if not len(sm_row): continue
    r2 = float(sm_row.iloc[0]['Vacancy_pct'])
    vac_by_sub_q[sm] = [safe(v*r2/all_r*100) for v in sav_q['VacPct']]

sav_m = sav_filt[sav_filt['Year'] >= 2018].sort_values(['Year','MN'])
sav_m['MLabel'] = sav_m['Year'].astype(str) + '-' + sav_m['MN'].astype(str).str.zfill(2)
vac_m_labels    = sav_m['MLabel'].tolist()
hki_vac_m       = [safe(v*hki_r/all_r*100) for v in sav_m['VacPct']]
kln_vac_m       = [safe(v*kln_r/all_r*100) for v in sav_m['VacPct']]
nt_vac_m        = [safe(v*nt_r/all_r*100)  for v in sav_m['VacPct']]
vac_by_sub_m    = {}
for sm in SUBMARKET_COLS.keys():
    sm_row = snap[snap['Location'] == sm]
    if not len(sm_row): continue
    r2 = float(sm_row.iloc[0]['Vacancy_pct'])
    vac_by_sub_m[sm] = [safe(v*r2/all_r*100) for v in sav_m['VacPct']]

# ── Vacant SF quarterly / monthly overall (SAV col-4) ─────────────────
sav_vsf_q = sav_vacsf[sav_vacsf['Year'] >= 2008].copy()
sav_vsf_q['Quarter'] = ((sav_vsf_q['MN'] - 1) // 3 + 1).astype(int)
sav_vsf_q = (sav_vsf_q.groupby(['Year','Quarter'], sort=False)
                       .apply(lambda x: x.nlargest(1,'MN'))
                       .reset_index(drop=True)
                       .sort_values(['Year','Quarter']))
sav_vsf_q['QLabel'] = sav_vsf_q['Year'].astype(str) + ' Q' + sav_vsf_q['Quarter'].astype(str)
vac_sf_q_labels     = sav_vsf_q['QLabel'].tolist()
vac_sf_overall_q    = [safe(v) for v in sav_vsf_q['VacSF']]

sav_vsf_m = sav_vacsf[sav_vacsf['Year'] >= 2018].sort_values(['Year','MN'])
sav_vsf_m['MLabel'] = sav_vsf_m['Year'].astype(str) + '-' + sav_vsf_m['MN'].astype(str).str.zfill(2)
vac_sf_m_labels     = sav_vsf_m['MLabel'].tolist()
vac_sf_overall_m    = [safe(v) for v in sav_vsf_m['VacSF']]

# ── Vacant SF quarterly / monthly per area/submarket/portfolio (Space) ─
df_space['Quarter'] = df_space['Date'].dt.quarter

# Quarterly: last date within each RefNo+Year+Quarter
sp_q = (df_space[df_space['Year'] >= 2008]
        .sort_values('Date')
        .drop_duplicates(['RefNo','Year','Quarter'], keep='last'))
sp_q['QLabel'] = sp_q['Year'].astype(str) + ' Q' + sp_q['Quarter'].astype(str)

def _sf_series_q(grp_col, grp_val, q_labels):
    sub = sp_q[sp_q[grp_col] == grp_val].groupby('QLabel')['Value'].sum()
    return [safe(sub.get(ql)) for ql in q_labels]

area_list_q = {a: _sf_series_q('Area', a, vac_sf_q_labels)
               for a in ['Hong Kong Island','Kowloon','New Territories']}
vac_sf_by_sub_q = {str(sm): _sf_series_q('Submarket', sm, vac_sf_q_labels)
                   for sm in sp_q['Submarket'].dropna().unique()}
vac_sf_by_port_q = {}
for pg_name in sp_q['PortGroup_s'].dropna().unique():
    if not pg_name: continue
    vac_sf_by_port_q[str(pg_name)] = _sf_series_q('PortGroup_s', pg_name, vac_sf_q_labels)

# Monthly: 2018+, last date within each RefNo+Year+Month
sp_m = (df_space[df_space['Year'] >= 2018]
        .sort_values('Date')
        .drop_duplicates(['RefNo','Year','Month'], keep='last'))
sp_m['MLabel'] = sp_m['Year'].astype(str) + '-' + sp_m['Month'].astype(str).str.zfill(2)

def _sf_series_m(grp_col, grp_val, m_labels):
    sub = sp_m[sp_m[grp_col] == grp_val].groupby('MLabel')['Value'].sum()
    return [safe(sub.get(ml)) for ml in m_labels]

area_list_m = {a: _sf_series_m('Area', a, vac_sf_m_labels)
               for a in ['Hong Kong Island','Kowloon','New Territories']}
vac_sf_by_sub_m = {str(sm): _sf_series_m('Submarket', sm, vac_sf_m_labels)
                   for sm in sp_m['Submarket'].dropna().unique()}
vac_sf_by_port_m = {}
for pg_name in sp_m['PortGroup_s'].dropna().unique():
    if not pg_name: continue
    vac_sf_by_port_m[str(pg_name)] = _sf_series_m('PortGroup_s', pg_name, vac_sf_m_labels)

# ── Portfolio NER & vacancy % quarterly / monthly (proportional scaling) ─
# Portfolio NER quarterly — Rental sheet actual data
df_rental_ner_pg['QLabel'] = (df_rental_ner_pg['Year'].astype(str) + ' Q' +
                               df_rental_ner_pg['Quarter'].astype(str))
_rner_q = (df_rental_ner_pg[df_rental_ner_pg['Year'] >= 2008]
           .groupby(['PortGroup', 'Year', 'Quarter', 'RefNo'])
           .apply(lambda x: x.nlargest(1, 'Month'))
           .reset_index(drop=True)
           .groupby(['PortGroup', 'QLabel'])['Value'].mean().reset_index())
ner_by_port_q = {}
for pg in portfolio_groups:
    nm = pg['name']
    sub = _rner_q[_rner_q['PortGroup'] == nm].set_index('QLabel')
    ner_by_port_q[nm] = [safe(sub.loc[ql, 'Value']) if ql in sub.index else None
                         for ql in ner_q_labels]

# Portfolio NER monthly — Rental sheet actual data (2018+)
df_rental_ner_pg['MLabel'] = df_rental_ner_pg['Date'].dt.strftime('%Y-%m')
_rner_m = (df_rental_ner_pg[df_rental_ner_pg['Year'] >= 2018]
           .groupby(['PortGroup', 'MLabel'])['Value'].mean().reset_index())
ner_by_port_m = {}
for pg in portfolio_groups:
    nm = pg['name']
    sub = _rner_m[_rner_m['PortGroup'] == nm].set_index('MLabel')
    ner_by_port_m[nm] = [safe(sub.loc[ml, 'Value']) if ml in sub.index else None
                         for ml in ner_m_labels]

# Portfolio Vacancy % quarterly/monthly — Space data / total NFA
vac_by_port_q, vac_by_port_m = {}, {}
for pg in portfolio_groups:
    nm = pg['name']
    total_nfa = port_nfa_map.get(nm, 0)
    if not total_nfa:
        continue
    sub_q = vac_sf_by_port_q.get(nm, [])
    vac_by_port_q[nm] = [safe(v / total_nfa * 100) if v is not None else None
                         for v in sub_q]
    sub_m = vac_sf_by_port_m.get(nm, [])
    vac_by_port_m[nm] = [safe(v / total_nfa * 100) if v is not None else None
                         for v in sub_m]

# District NER quarterly (Rental sheet)
df_rental_ner_dist['QLabel'] = (df_rental_ner_dist['Year'].astype(str) + ' Q' +
                                df_rental_ner_dist['Quarter'].astype(str))
_rner_dist_q = (df_rental_ner_dist[df_rental_ner_dist['Year'] >= 2008]
                .groupby(['District', 'Year', 'Quarter', 'RefNo'])
                .apply(lambda x: x.nlargest(1, 'Month'))
                .reset_index(drop=True)
                .groupby(['District', 'QLabel'])['Value'].mean().reset_index())
ner_by_dist_q = {}
for d in district_snapshot:
    nm = d['name']
    sub = _rner_dist_q[_rner_dist_q['District'] == nm].set_index('QLabel')
    ner_by_dist_q[nm] = [safe(sub.loc[ql, 'Value']) if ql in sub.index else None for ql in ner_q_labels]

# District NER monthly (Rental sheet, 2018+)
df_rental_ner_dist['MLabel'] = df_rental_ner_dist['Date'].dt.strftime('%Y-%m')
_rner_dist_m = (df_rental_ner_dist[df_rental_ner_dist['Year'] >= 2018]
                .groupby(['District', 'MLabel'])['Value'].mean().reset_index())
ner_by_dist_m = {}
for d in district_snapshot:
    nm = d['name']
    sub = _rner_dist_m[_rner_dist_m['District'] == nm].set_index('MLabel')
    ner_by_dist_m[nm] = [safe(sub.loc[ml, 'Value']) if ml in sub.index else None for ml in ner_m_labels]

# District vacant SF quarterly/monthly (Space sheet)
def _sf_series_q_dist(dist_name, q_labels):
    sub = sp_q[sp_q['District_s'] == dist_name].groupby('QLabel')['Value'].sum()
    return [safe(sub.get(ql)) for ql in q_labels]
def _sf_series_m_dist(dist_name, m_labels):
    sub = sp_m[sp_m['District_s'] == dist_name].groupby('MLabel')['Value'].sum()
    return [safe(sub.get(ml)) for ml in m_labels]

vac_sf_by_dist_q = {str(d): _sf_series_q_dist(d, vac_sf_q_labels)
                    for d in sp_q['District_s'].dropna().unique() if d}
vac_sf_by_dist_m = {str(d): _sf_series_m_dist(d, vac_sf_m_labels)
                    for d in sp_m['District_s'].dropna().unique() if d}

# District vacancy % quarterly/monthly from Space + NFA
vac_by_dist_q, vac_by_dist_m = {}, {}
for d in district_snapshot:
    nm = d['name']
    total_nfa = dist_nfa_map.get(nm, 0)
    if not total_nfa: continue
    sub_q = vac_sf_by_dist_q.get(nm, [])
    vac_by_dist_q[nm] = [safe(v / total_nfa * 100) if v is not None else None for v in sub_q]
    sub_m = vac_sf_by_dist_m.get(nm, [])
    vac_by_dist_m[nm] = [safe(v / total_nfa * 100) if v is not None else None for v in sub_m]

print(f"    Q series: {len(ner_q_labels)} quarters | M series: {len(ner_m_labels)} months (2018+)")

# ══════════════════════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════════════════════
# NET ABSORPTION TIME SERIES
# All series from Output sheet NA columns (cols 43-81) — monthly, all locations.
# Quarterly = sum of 3 monthly values (NA is a flow, not a stock).
# Annual    = sum of 12 monthly values.
# All 12 months including September now available in the database.
# ══════════════════════════════════════════════════════════════════════

# na_out was built alongside nr earlier.  Filter to 2008+.
na_filt = na_out[na_out['Year'] >= 2008].copy()
na_filt['QLabel'] = na_filt['Year'].astype(str) + ' Q' + na_filt['Quarter'].astype(str)
na_filt['MLabel'] = na_filt['Date'].dt.strftime('%Y-%m')

# Collect all non-metadata columns for aggregation
_na_data_cols = [c for c in na_filt.columns
                 if c not in ('Date','Year','Month','Quarter','QLabel','MLabel')]
_dist_snap_names = {d['name'] for d in district_snapshot}

# ── Monthly (2018+) ───────────────────────────────────────────────────
na_m_filt    = na_filt[na_filt['Year'] >= 2018].sort_values(['Year','Month'])
na_m_labels  = na_m_filt['MLabel'].tolist()
na_overall_m = [safe(v) for v in na_m_filt['Overall Hong Kong']]
na_hki_m     = [safe(v) for v in na_m_filt['Hong Kong Island']]
na_kln_m     = [safe(v) for v in na_m_filt['Kowloon']]
na_nt_m      = [safe(v) for v in na_m_filt['New Territories']]

na_by_sub_m  = {}
for sm in SUBMARKET_COLS:
    if sm in na_m_filt.columns:
        series = [safe(v) for v in na_m_filt[sm]]
        if any(v is not None for v in series):
            na_by_sub_m[sm] = series

na_by_dist_m = {}
for col in _na_data_cols:
    if col in _dist_snap_names:
        series = [safe(v) for v in na_m_filt[col]] if col in na_m_filt.columns else []
        if any(v is not None for v in series):
            na_by_dist_m[col] = series

# ── Quarterly (2008+) — sum monthly values per quarter ───────────────
na_q = (na_filt.groupby(['Year','Quarter'])[_na_data_cols]
               .sum(min_count=1).reset_index())
na_q['QLabel'] = na_q['Year'].astype(str) + ' Q' + na_q['Quarter'].astype(str)
na_q = na_q.sort_values(['Year','Quarter'])

na_q_labels  = na_q['QLabel'].tolist()
na_overall_q = [safe(v) for v in na_q['Overall Hong Kong']]
na_hki_q     = [safe(v) for v in na_q['Hong Kong Island']]
na_kln_q     = [safe(v) for v in na_q['Kowloon']]
na_nt_q      = [safe(v) for v in na_q['New Territories']]

na_by_sub_q = {}
for sm in SUBMARKET_COLS:
    if sm in na_q.columns:
        series = [safe(v) for v in na_q[sm]]
        if any(v is not None for v in series):
            na_by_sub_q[sm] = series

na_by_dist_q = {}
for col in _na_data_cols:
    if col in _dist_snap_names and col in na_q.columns:
        series = [safe(v) for v in na_q[col]]
        if any(v is not None for v in series):
            na_by_dist_q[col] = series

# ── Annual (2008+) — sum monthly values per year ─────────────────────
na_ann = (na_filt.groupby('Year')[_na_data_cols]
                 .sum(min_count=1).reset_index()
                 .sort_values('Year'))

na_ann_labels  = na_ann['Year'].tolist()
na_overall_ann = [safe(v) for v in na_ann['Overall Hong Kong']]
na_hki_ann     = [safe(v) for v in na_ann['Hong Kong Island']]
na_kln_ann     = [safe(v) for v in na_ann['Kowloon']]
na_nt_ann      = [safe(v) for v in na_ann['New Territories']]

na_by_submarket = {}
for sm in SUBMARKET_COLS:
    if sm in na_ann.columns:
        series = [safe(v) for v in na_ann[sm]]
        if any(v is not None for v in series):
            na_by_submarket[sm] = series

na_by_district = {}
for col in _na_data_cols:
    if col in _dist_snap_names and col in na_ann.columns:
        series = [safe(v) for v in na_ann[col]]
        if any(v is not None for v in series):
            na_by_district[col] = series

print(f"    Net Absorption: {len(na_q_labels)} quarters, {len(na_ann_labels)} annual years, "
      f"{len(na_by_sub_q)} submarkets Q | {len(na_by_sub_m)} submarkets M | "
      f"{len(na_by_district)} districts")

bldg_records = []
for _, r in bldg_merged.iterrows():
    if pd.notna(r['Lat']) and pd.notna(r['Long']) and r['Lat']!=0:
        bldg_records.append({
            'name': str(r['BldgName']),
            'district': str(r['District']) if pd.notna(r['District']) else '',
            'submarket': str(r['Submarket']) if pd.notna(r['Submarket']) else '',
            'area': str(r['Area']) if pd.notna(r['Area']) else '',
            'grade': str(r['Grade']) if pd.notna(r['Grade']) else '',
            'ownership': str(r['Ownership']) if pd.notna(r['Ownership']) else '',
            'landlord': str(r['Landlord']) if pd.notna(r['Landlord']) else '',
            'compYear': safe(r['CompYear']),
            'nfa': safe(r['NFA']),
            'flrPlateNFA': safe(r['FlrPlateNFA']) if 'FlrPlateNFA' in r.index and pd.notna(r['FlrPlateNFA']) else None,
            'ner': safe(r['NER']),
            'vacant': safe(r['Vacant']),
            'futureVac': safe(r['FutureVac']) if 'FutureVac' in r.index and pd.notna(r['FutureVac']) else None,
            'shadowVac': safe(r['ShadowVac']) if 'ShadowVac' in r.index and pd.notna(r['ShadowVac']) else None,
            'vacPct': safe(r['VacPct']*100) if pd.notna(r['VacPct']) else None,
            'lat': safe(r['Lat']),
            'lon': safe(r['Long']),
        })

data['rent_vacancy'] = {
    'snap_date': snap_date,
    'overall_ner': safe(overall_row['NER']),
    'overall_vac': safe(overall_row['Vacancy_pct']*100),
    'vacancy_by_district': {'labels': snap_vac['Location'].tolist(), 'values': snap_vac['vac'].tolist()},
    'ytd_rent_change':     {'labels': snap_ytd['Location'].tolist(), 'values': snap_ytd['ytd'].tolist()},
    'rent_change_by_loc':  rent_change_by_loc,
    'ner_trend': {'years': ner_yrs, 'hki': hki_ner, 'kowloon': kln_ner, 'nt': nt_ner},
    'ner_by_submarket': ner_by_submarket,      # per-submarket NER time series
    'vac_by_submarket': vac_by_submarket,      # per-submarket vacancy % time series
    'vac_years_sm':     vac_yrs_sm,            # year axis for vac_by_submarket
    'submarket_snapshot': submarket_snapshot,  # current NER + vacancy per submarket
    'portfolio_groups':   portfolio_groups,    # current NER + vacancy per portfolio group
    'ner_by_portfolio':   ner_by_portfolio,    # per-portfolio NER time series (scaled)
    'vac_by_portfolio':   vac_by_portfolio,    # per-portfolio vacancy time series (scaled)
    'vac_trend': {'years': vac_yrs, 'hki': hki_vac, 'kowloon': kln_vac, 'nt': nt_vac},
    # Vacant Space (sq.ft. NFA) time series — from Space sheet (not converted from %)
    'vac_sf_trend': {
        'years': area_years,
        'overall': vac_sf_overall,
        'hki': vac_sf_by_area.get('Hong Kong Island', []),
        'kowloon': vac_sf_by_area.get('Kowloon', []),
        'nt': vac_sf_by_area.get('New Territories', []),
    },
    'vac_sf_by_submarket': vac_sf_by_submarket,  # per-submarket vacant sf time series (annual)
    'vac_sf_by_portfolio': vac_sf_by_portfolio,  # per-portfolio vacant sf time series (annual)
    # ── Quarterly series ('YYYY QN', 2008+) ──────────────────────────────
    'ner_trend_q':     {'labels': ner_q_labels, 'hki': hki_ner_q, 'kowloon': kln_ner_q, 'nt': nt_ner_q},
    'vac_trend_q':     {'labels': vac_q_labels, 'hki': hki_vac_q, 'kowloon': kln_vac_q, 'nt': nt_vac_q},
    'vac_sf_trend_q':  {'labels': vac_sf_q_labels, 'overall': vac_sf_overall_q,
                         'hki': area_list_q.get('Hong Kong Island',[]),
                         'kowloon': area_list_q.get('Kowloon',[]),
                         'nt': area_list_q.get('New Territories',[])},
    'ner_by_sub_q':    ner_by_sub_q,
    'vac_by_sub_q':    vac_by_sub_q,
    'vac_sf_by_sub_q': vac_sf_by_sub_q,
    'ner_by_port_q':   ner_by_port_q,
    'vac_by_port_q':   vac_by_port_q,
    'vac_sf_by_port_q':vac_sf_by_port_q,
    # ── Monthly series ('YYYY-MM', 2018+) ────────────────────────────────
    'ner_trend_m':     {'labels': ner_m_labels, 'hki': hki_ner_m, 'kowloon': kln_ner_m, 'nt': nt_ner_m},
    'vac_trend_m':     {'labels': vac_m_labels, 'hki': hki_vac_m, 'kowloon': kln_vac_m, 'nt': nt_vac_m},
    'vac_sf_trend_m':  {'labels': vac_sf_m_labels, 'overall': vac_sf_overall_m,
                         'hki': area_list_m.get('Hong Kong Island',[]),
                         'kowloon': area_list_m.get('Kowloon',[]),
                         'nt': area_list_m.get('New Territories',[])},
    'ner_by_sub_m':    ner_by_sub_m,
    'vac_by_sub_m':    vac_by_sub_m,
    'vac_sf_by_sub_m': vac_sf_by_sub_m,
    'ner_by_port_m':   ner_by_port_m,
    'vac_by_port_m':   vac_by_port_m,
    'vac_sf_by_port_m':vac_sf_by_port_m,
    # ── Net Absorption (quarterly source data, annual = sum of quarters) ─
    'district_snapshot':     district_snapshot,
    'ner_by_district':        ner_by_district,
    'vac_by_district':        vac_by_district,
    'vac_sf_by_district':     vac_sf_by_district,
    'ner_by_dist_q':          ner_by_dist_q,
    'vac_by_dist_q':          vac_by_dist_q,
    'vac_sf_by_dist_q':       vac_sf_by_dist_q,
    'ner_by_dist_m':          ner_by_dist_m,
    'vac_by_dist_m':          vac_by_dist_m,
    'vac_sf_by_dist_m':       vac_sf_by_dist_m,
    'na_trend':        {'years': na_ann_labels, 'overall': na_overall_ann,
                         'hki': na_hki_ann, 'kowloon': na_kln_ann, 'nt': na_nt_ann},
    'na_trend_q':      {'labels': na_q_labels, 'overall': na_overall_q,
                         'hki': na_hki_q,  'kowloon': na_kln_q,  'nt': na_nt_q},
    'na_trend_m':      {'labels': na_m_labels, 'overall': na_overall_m},
    'na_by_submarket': na_by_submarket,  # annual per-submarket net absorption
    'na_by_sub_q':     na_by_sub_q,      # quarterly per-submarket (Output + Q3 from Space)
    'na_by_sub_m':     na_by_sub_m,      # monthly per-submarket (Space sheet, 2018+)
    'na_by_district':  na_by_district,   # annual per-district NA
    'na_by_dist_q':    na_by_dist_q,     # quarterly per-district (Q3 filled from Space)
    'na_by_dist_m':    na_by_dist_m,     # monthly per-district (Space sheet, 2018+)
    'top_vacant': {'labels': top_vac['Name'].tolist(), 'values': [safe(v/1000) for v in top_vac['Vacant']]},
    'top_ner':    {'labels': top_ner['Name'].tolist(), 'values': [safe(v) for v in top_ner['NER']]},
    'buildings':  bldg_records,
}
print(f"    Rent & Vacancy: {len(bldg_records)} buildings, snap={snap_date}")

# ══════════════════════════════════════════════════════════════════════
# 2. TRANSACTIONS — Leasing Transactions 2.0.xlsx
# ══════════════════════════════════════════════════════════════════════
print("  Reading Leasing Transactions...")
df_tx = pd.read_excel(os.path.join(DB, "Leasing Transactions 2.0.xlsx"),
                      sheet_name='Transactions', header=None)
# Cols: 0=Date,1=Year,2=Month,3=Quarter,5=Area,6=District,7=Submarket,
#       8=Building,9=Grade(detailed),10=Grade2(A/B),13=Tenant,15=TenantRegion,17=Sector,19=TxnType,46=GFA,47=NFA,59=Agency,70=BuildingCode
txn = df_tx.iloc[1:, [0,1,2,3,5,6,7,8,9,10,13,15,17,19,46,47,59,70]].copy()
txn.columns = ['Date','Year','Month','Quarter','Area','District','Submarket',
               'Building','Grade','Grade2','Tenant','TenantRegion','Sector','TxnType','GFA','NFA','Agency','BuildingCode']
txn['Agency'] = txn['Agency'].astype(str).str.strip()
txn['Agency'] = txn['Agency'].replace({'nan':'','Direct ':'Direct','JLL ':'JLL','Savills ':'Savills',
                                        'Blackhill':'Blackhills','BlackHills':'Blackhills',
                                        'Local agent':'Local Agent','local agent':'Local Agent',
                                        'direct':'Direct','Treasure land':'Treasureland'})
txn['Date'] = pd.to_datetime(txn['Date'], errors='coerce')
txn['Year'] = pd.to_numeric(txn['Year'], errors='coerce')
txn['Month'] = pd.to_numeric(txn['Month'], errors='coerce')
txn['GFA'] = pd.to_numeric(txn['GFA'], errors='coerce')
txn['NFA'] = pd.to_numeric(txn['NFA'], errors='coerce')
txn = txn.dropna(subset=['Date','Year'])

# ── Apply Bldg List categorisation (District / Submarket / Area) ──────
# For each transaction, look up its BuildingCode in bldg_map; if matched
# override the three geography fields with Bldg List values (source of truth).
def apply_bldg_map(row):
    code = str(row['BuildingCode']) if pd.notna(row['BuildingCode']) else ''
    info = bldg_map.get(code)
    if info:
        row['District']  = info['district']  or row['District']
        row['Submarket'] = info['submarket'] or row['Submarket']
        row['Area']      = info['area']      or row['Area']
    return row

txn = txn.apply(apply_bldg_map, axis=1)

# Normalize TxnType (handle mixed case entries like "new Letting", "renewal")
txn['TxnType_norm'] = txn['TxnType'].astype(str).str.strip().str.title()

# ── Grade A + New Letting filter ──────────────────────────────────────
# Grade2 col = 'A' covers A1/A2/A3; also catch Grade col directly
gradeA_mask = txn['Grade2'].astype(str).str.strip().str.upper() == 'A'
nl_mask     = txn['TxnType_norm'] == 'New Letting'
gradeA_nl   = txn[gradeA_mask & nl_mask].copy()
gradeA_nl['YM'] = gradeA_nl['Date'].dt.to_period('M').astype(str)

# Extract quarter number (1-4) from Quarter column like "2007 Q1"
def parse_quarter(q):
    try: return int(str(q).strip()[-1])
    except: return None
gradeA_nl['Qnum'] = gradeA_nl['Quarter'].apply(parse_quarter)

# ── Annual volumes (Grade A NL, 2007+) ────────────────────────────────
ann_txn = gradeA_nl.groupby('Year').agg(Volume=('NFA','sum'), Count=('Date','count')).reset_index()
ann_txn = ann_txn[ann_txn['Year']>=2007].sort_values('Year')

# ── Monthly rolling 24 months (Grade A NL) ────────────────────────────
latest_ym = gradeA_nl['YM'].max()
cutoff    = str(pd.Period(latest_ym,'M') - 23)
monthly_txn = gradeA_nl.groupby('YM').agg(Volume=('NFA','sum'), Count=('Date','count')).reset_index()
monthly_txn = monthly_txn[monthly_txn['YM']>=cutoff].sort_values('YM')
monthly_labels = monthly_txn['YM'].tolist()   # fixed 24-month window for JS

# ── Sector breakdown (Grade A NL, all time) ────────────────────────────
sector_nl    = gradeA_nl.groupby('Sector')['NFA'].sum().sort_values(ascending=False)
sector_total = sector_nl.sum()
sector_pct   = (sector_nl/sector_total*100).round(1).head(10)

# ── Compact raw dataset for dynamic JS filtering (Grade A NL, 2007+) ──
# Fields: y=year, q=quarter, ym=YYYY-MM, sub=submarket, area=HKI/Kowloon/NT,
#         sec=sector, nfa=sf
gradeA_nl_2007 = gradeA_nl[gradeA_nl['Year']>=2007]
raw_nl = []
for _, r in gradeA_nl_2007.iterrows():
    nfa = r['NFA']
    if pd.isna(nfa) or nfa <= 0: continue
    raw_nl.append({
        'y':   int(r['Year']),
        'q':   r['Qnum'],
        'ym':  str(r['YM']),
        'sub': str(r['Submarket']) if pd.notna(r['Submarket']) else '',
        'area':str(r['Area'])         if pd.notna(r['Area'])         else '',
        'sec': str(r['Sector'])       if pd.notna(r['Sector'])       else '',
        'rgn': str(r['TenantRegion']) if pd.notna(r['TenantRegion']) and str(r['TenantRegion']).strip() not in ('','nan','Unknown') else 'Unknown',
        'agn': str(r['Agency'])       if pd.notna(r['Agency'])       and str(r['Agency']).strip()       not in ('','nan') else '',
        'nfa': round(float(nfa), 0),
    })

# ── Unique filter option lists ─────────────────────────────────────────
# Submarket list comes from Bldg List (source of truth), not transaction file
unique_submkts   = sorted(bldg_info['Submarket'].dropna().unique().tolist())
unique_sectors   = sorted(gradeA_nl['Sector'].dropna().unique().tolist())
unique_years     = sorted(int(y) for y in ann_txn['Year'].tolist())
unique_agencies  = sorted(a for a in gradeA_nl['Agency'].dropna().unique() if str(a).strip() not in ('','nan'))

# ── Grade A New Lettings 2026 for table (consistent with charts) ──────
deals_2026 = txn[gradeA_mask & nl_mask & (txn['Year'] == 2026)].sort_values('Date', ascending=False)

data['transactions'] = {
    'latest_date':  gradeA_nl['Date'].max().strftime('%b %Y'),
    'total_count':  int(len(txn)),            # all deals in database
    'gradeA_nl_count': int(len(gradeA_nl)),   # Grade A NL deals
    'annual': {
        'years':  ann_txn['Year'].tolist(),
        'volume': [safe(v/1e6) for v in ann_txn['Volume']],
        'count':  ann_txn['Count'].tolist(),
    },
    'monthly': {
        'labels': monthly_labels,
        'volume': [safe(v/1e3) for v in monthly_txn['Volume']],
        'count':  monthly_txn['Count'].tolist(),
    },
    'sector': {
        'labels': sector_pct.index.tolist(),
        'pct':    sector_pct.tolist(),
    },
    # Compact raw records for client-side filtering
    'raw_nl': raw_nl,
    'submkts':  unique_submkts,
    'sectors':  unique_sectors,
    'years':    unique_years,
    'agencies': unique_agencies,
    'monthly_labels': monthly_labels,    # 24-month window for JS charts
    'latest_deals': [
        {
            'date':     r['Date'].strftime('%Y-%m-%d'),
            'building': str(r['Building'])  if pd.notna(r['Building'])  else '',
            'submarket':str(r['Submarket']) if pd.notna(r['Submarket']) else '',
            'district': str(r['District'])  if pd.notna(r['District'])  else '',
            'area':     str(r['Area'])      if pd.notna(r['Area'])      else '',
            'tenant':   str(r['Tenant'])    if pd.notna(r['Tenant'])    else '',
            'sector':   str(r['Sector'])    if pd.notna(r['Sector'])    else '',
            'type':     str(r['TxnType'])   if pd.notna(r['TxnType'])   else '',
            'nfa':      safe(r['NFA']),
            'gfa':      safe(r['GFA']),
        }
        for _, r in deals_2026.iterrows()
    ],
    # All-time Grade A NL deals for tenant/building/agency search (2007+), sorted newest first
    'all_nl_deals': [
        {
            'date':     r['Date'].strftime('%Y-%m-%d'),
            'building': str(r['Building'])  if pd.notna(r['Building'])  else '',
            'submarket':str(r['Submarket']) if pd.notna(r['Submarket']) else '',
            'district': str(r['District'])  if pd.notna(r['District'])  else '',
            'tenant':   str(r['Tenant'])    if pd.notna(r['Tenant'])    else '',
            'sector':   str(r['Sector'])    if pd.notna(r['Sector'])    else '',
            'nfa':      safe(r['NFA']),
            'agn':      str(r['Agency'])    if pd.notna(r['Agency']) and str(r['Agency']).strip() not in ('','nan') else '',
            'rgn':      str(r['TenantRegion']) if pd.notna(r['TenantRegion']) and str(r['TenantRegion']).strip() not in ('','nan','Unknown') else 'Unknown',
        }
        for _, r in gradeA_nl_2007.sort_values('Date', ascending=False).iterrows()
        if pd.notna(r['Tenant']) and str(r['Tenant']).strip()
    ],
    # Sorted unique tenant list for autocomplete
    'all_tenants': sorted(
        t for t in gradeA_nl['Tenant'].dropna().unique()
        if str(t).strip() and str(t).lower() not in ('undisclosed tenant','undisclosed','nan')
    ),
    # Sorted unique building list for autocomplete
    'all_buildings': sorted(
        b for b in gradeA_nl['Building'].dropna().unique()
        if str(b).strip() and str(b).lower() not in ('','nan')
    ),
}
print(f"    Transactions: {len(txn)} total, {len(gradeA_nl)} Grade A NL, {len(raw_nl)} raw NL records (2015+)")

# ══════════════════════════════════════════════════════════════════════
# 3. FUTURE SUPPLY — Development Pipeline 2.1.xlsx  ('For Charts' sheet)
# ══════════════════════════════════════════════════════════════════════
print("  Reading Development Pipeline...")
# Source: 'Grade A Office Supply_pre-2020' sheet — full project-level data
# Filter: Type='A Office', Status='Valid', Year 2026–2032
# NFA = col 24 'Office UFA (Net) (sf)'
df_pipeline = pd.read_excel(os.path.join(DB, "Development Pipeline 2.1.xlsx"),
                            sheet_name='Grade A Office Supply_pre-2020', header=0)
df_pipeline['Year'] = pd.to_numeric(df_pipeline['Year'], errors='coerce')
df_pipeline['NFA']  = pd.to_numeric(df_pipeline['Office UFA (Net) (sf)'], errors='coerce')

fc_right = df_pipeline[
    (df_pipeline['Type'] == 'A Office') &
    (df_pipeline['Status'] == 'Valid') &
    (df_pipeline['Year'] >= 2026) &
    (df_pipeline['Year'] <= 2032)
].copy()
fc_right = fc_right.dropna(subset=['Year','Building Name','NFA'])
fc_right['Year'] = fc_right['Year'].astype(int)
# Rename columns to match downstream code
fc_right = fc_right.rename(columns={
    'Building Name':     'Building',
    'CBRE districts':    'District',
    'CBRE sub-markets':  'Submarket',
    'Developer':         'Developer',
    'Item Code':         'ItemCode',
})

# Map districts to HKI / Kowloon / NT
hki_dists = {'Central','Admiralty & Sheung Wan','Wan Chai & Causeway Bay',
             'Hong Kong East','Wong Chuk Hang','HKI Others'}
kln_dists = {'Tsim Sha Tsui','Tsim Sha Tsui West','Kowloon East','Kowloon Others',
             'Decentralised Kowloon'}
def area_group(d):
    if str(d) in hki_dists: return 'Hong Kong Island'
    if str(d) in kln_dists: return 'Kowloon'
    return 'New Territories'
fc_right['AreaGroup'] = fc_right['District'].apply(area_group)

# Pad full 2026–2032 range so all 7 years always appear
SUPPLY_RANGE = list(range(2026, 2033))

yr_grp = fc_right.groupby(['Year','AreaGroup'])['NFA'].sum().unstack(fill_value=0)
supply_years = SUPPLY_RANGE
supply_hki  = [int(yr_grp.loc[y,'Hong Kong Island']) if y in yr_grp.index and 'Hong Kong Island' in yr_grp.columns else 0 for y in SUPPLY_RANGE]
supply_kln  = [int(yr_grp.loc[y,'Kowloon'])          if y in yr_grp.index and 'Kowloon'          in yr_grp.columns else 0 for y in SUPPLY_RANGE]
supply_nt   = [int(yr_grp.loc[y,'New Territories'])  if y in yr_grp.index and 'New Territories'  in yr_grp.columns else 0 for y in SUPPLY_RANGE]
supply_total= [h+k+n for h,k,n in zip(supply_hki, supply_kln, supply_nt)]

# District donut (2026–2032 only)
dist_grp = fc_right.groupby('District')['NFA'].sum().sort_values(ascending=False)
dist_donut = {'labels': dist_grp.index.tolist(), 'values': [int(v) for v in dist_grp.values]}

# Project table (2026–2032 only)
proj_list = []
for _, r in fc_right.iterrows():
    proj_list.append({
        'building':  str(r['Building']),
        'district':  str(r['District']) if pd.notna(r['District']) else '',
        'submarket': str(r['Submarket']) if pd.notna(r['Submarket']) else '',
        'developer': str(r['Developer']) if pd.notna(r['Developer']) else '',
        'year':      int(r['Year']),
        'nfa_sf':    safe(r['NFA']),
        'area':      str(r['AreaGroup']),
        'lat':       float(r['Lat'])  if pd.notna(r['Lat'])  else None,
        'lng':       float(r['Long']) if pd.notna(r['Long']) else None,
    })

total_nfa_m = fc_right['NFA'].sum() / 1e6
data['future_supply'] = {
    'total_projects': len(proj_list),
    'total_nfa': f'{total_nfa_m:.1f}M',
    'supply_by_year': {
        'years': supply_years,
        'hki': supply_hki,
        'kowloon': supply_kln,
        'nt': supply_nt,
        'total': supply_total,
    },
    'supply_by_district': dist_donut,
    'projects': proj_list,
}
print(f"    Future Supply (2026–2032): {len(proj_list)} projects, {total_nfa_m:.1f}M sf NFA")

# ══════════════════════════════════════════════════════════════════════
# WRITE OUTPUT
# ══════════════════════════════════════════════════════════════════════
data['generated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M')
with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(data, f, indent=2, ensure_ascii=False, default=str)

# Also write a JS embed file for local file:// access (no CORS issues)
js_path = os.path.join(BASE, "dashboard_data.js")
with open(js_path, 'w', encoding='utf-8') as f:
    f.write("window.DASHBOARD_DATA = ")
    json.dump(data, f, indent=2, ensure_ascii=False, default=str)
    f.write(";")

size_kb = os.path.getsize(OUT) / 1024
print(f"\n✓ dashboard_data.json written ({size_kb:.0f} KB) at {data['generated_at']}")
print(f"  Buildings: {len(data['rent_vacancy']['buildings'])}")
print(f"  Transactions: {data['transactions']['total_count']}")
print(f"  Supply projects: {data['future_supply']['total_projects']}")

# ══════════════════════════════════════════════════════════════════════
# STACKING PLAN DATA  (generates stacking_plan_data.js for Occupancy tab)
# ══════════════════════════════════════════════════════════════════════
_sp_out_for_enc = None   # captured below for GitHub encryption
try:
    print("\nBuilding stacking plan data...")
    sp_path  = os.path.join(DB, "Grade A Office Stacking Plan.xlsx")
    ref_path = os.path.join(DB, "Reference List.xlsx")

    sp  = pd.read_excel(sp_path,  sheet_name='ft_STACKING_PLAN', header=0)
    ref = pd.read_excel(ref_path)

    sp['GFA']  = pd.to_numeric(sp['GFA'], errors='coerce')
    sp['Date'] = pd.to_datetime(sp['Date'], errors='coerce')
    sp = sp.dropna(subset=['Date'])

    ref_map = (ref[['Building Code','District','Sub-market','Efficiency (%)']].copy()
               .drop_duplicates('Building Code'))
    ref_map['Efficiency (%)'] = pd.to_numeric(ref_map['Efficiency (%)'], errors='coerce').fillna(0.75)
    ref_map['Sub-market']     = ref_map['Sub-market'].str.lstrip('0').str.strip()
    ref_map['District']       = ref_map['District'].str.lstrip('0').str.strip()

    # Sector and Origin are now embedded in the stacking plan file — no TNM merge needed
    df = sp.merge(ref_map, on='Building Code', how='left')
    df['Efficiency (%)'] = df['Efficiency (%)'].fillna(0.75)
    df['NFA'] = df['GFA'] * df['Efficiency (%)']
    df['YM']  = df['Date'].dt.to_period('M').astype(str)

    # Occupied = Status != 'For Lease'  AND  Tenant/Sector/Origin/Region not literally 'vacant'
    for_lease_mask = df['Status'].astype(str).str.lower() == 'for lease'
    vacant_mask = (
        (df['Tenant'].astype(str).str.lower() == 'vacant') |
        (df['Tenant Sector'].astype(str).str.lower() == 'vacant') |
        (df['Tenant Origin'].astype(str).str.lower() == 'vacant') |
        (df['Tenant Region'].astype(str).str.lower() == 'vacant')
    )
    occ = df[~for_lease_mask & ~vacant_mask].copy()
    occ['Tenant Sector'] = occ['Tenant Sector'].fillna('Unidentified')
    occ['Tenant Region'] = occ['Tenant Region'].fillna('Unknown')
    occ['District']      = occ['District'].fillna('Unknown')

    # Periods: one snapshot per unique (year, month) present in the full dataset
    periods = sorted(df['YM'].unique())

    # ── Lease Expiry chart data (latest period only) ─────────────────────
    EXPIRY_YEARS = [2026, 2027, 2028, 2029]
    latest_p     = sorted(df['YM'].unique())[-1]
    latest_all   = df[df['YM'] == latest_p].copy()
    latest_all['ExpiryDate'] = pd.to_datetime(latest_all['Expiry Date'], errors='coerce')
    latest_all['ExpiryYear'] = latest_all['ExpiryDate'].dt.year
    latest_all['District']   = latest_all['District'].fillna('Unknown')

    SHADOW_STATUSES  = {'Replace', 'Sublet', 'Surrender'}
    FUTVAC_STATUSES  = {'FutureVac'}

    # Three mutually exclusive slices (all restricted to expiry years 2026-2029)
    in_years = latest_all['ExpiryYear'].isin(EXPIRY_YEARS)

    # 1. Lease Expiry — excl future vacancy, shadow space, and owner-occupied
    mask_le = (
        ~latest_all['Status'].astype(str).isin(FUTVAC_STATUSES | SHADOW_STATUSES) &
        (latest_all['Possession'].astype(str) != 'Owner Occupies')
    )
    # 2. Future Vacancy
    mask_fv = latest_all['Status'].astype(str).isin(FUTVAC_STATUSES)
    # 3. Shadow Space
    mask_ss = latest_all['Status'].astype(str).isin(SHADOW_STATUSES)

    def _expiry_by_dist(mask):
        base = latest_all[mask & in_years].copy()
        dists = sorted(base['District'].unique().tolist())
        return {d: [round(float(base[base['District']==d][base['ExpiryYear']==y]['NFA'].sum()), 0)
                    for y in EXPIRY_YEARS] for d in dists}

    # Union of districts across all three slices (for consistent filter list)
    all_dists = sorted(set(
        latest_all[in_years & (mask_le | mask_fv | mask_ss)]['District'].unique()
    ))

    def _expiry_by_dist2(mask):
        base = latest_all[mask & in_years]
        result = {}
        for d in all_dists:
            result[d] = [round(float(base[base['District']==d][base['ExpiryYear']==y]['NFA'].sum()), 0)
                         for y in EXPIRY_YEARS]
        return result

    expiry_data = {
        'years':          EXPIRY_YEARS,
        'districts':      all_dists,
        'lease_expiry':   _expiry_by_dist2(mask_le),
        'future_vacancy': _expiry_by_dist2(mask_fv),
        'shadow_space':   _expiry_by_dist2(mask_ss),
    }

    # ── Cross-aggregation for occupied space filter (sector × district × region) ──
    cross_sectors  = sorted(occ['Tenant Sector'].unique().tolist())
    cross_dists    = sorted(occ['District'].unique().tolist())
    cross_regions  = sorted(occ['Tenant Region'].unique().tolist())
    sec_idx  = {s: i for i, s in enumerate(cross_sectors)}
    dist_idx = {d: i for i, d in enumerate(cross_dists)}
    reg_idx  = {r: i for i, r in enumerate(cross_regions)}

    cross_by_period = {}
    for p in periods:
        sub = occ[occ['YM'] == p]
        grp = (sub.groupby(['Tenant Sector', 'District', 'Tenant Region'])['NFA']
                  .sum().reset_index())
        grp = grp[grp['NFA'] > 0]
        cross_by_period[p] = [
            [sec_idx[r['Tenant Sector']], dist_idx[r['District']],
             reg_idx[r['Tenant Region']], round(float(r['NFA']), 0)]
            for _, r in grp.iterrows()
        ]

    cross_data = {
        'sectors':   cross_sectors,
        'districts': cross_dists,
        'regions':   cross_regions,
        'by_period': cross_by_period,
    }

    sp_out = {'periods': periods, 'by_sector': {}, 'by_submarket': {}, 'by_origin': {}, 'summary': {},
              'expiry': expiry_data, 'cross': cross_data}

    for p in periods:
        # All units in this period (for total stock)
        all_df  = df[df['YM'] == p]
        sub_df  = occ[occ['YM'] == p]
        sp_out['summary'][p] = {
            'buildings':    int(all_df['Building Code'].nunique()),
            'tenants':      int(sub_df['Tenant'].nunique()),
            'total_nfa':    round(float(all_df['NFA'].sum()), 0),
            'occupied_nfa': round(float(sub_df['NFA'].sum()), 0),
            'premises':     int(len(all_df)),
        }
        # by_submarket → District (displayed as "Submarket" in dashboard)
        # by_origin    → Tenant Region (displayed as "Origin" in dashboard)
        for dim_key, col in [('by_sector','Tenant Sector'), ('by_submarket','District'), ('by_origin','Tenant Region')]:
            grp = sub_df.groupby(col)['NFA'].sum().sort_values(ascending=False)
            sp_out[dim_key][p] = {
                'labels': grp.index.tolist(),
                'values': [round(v, 0) for v in grp.values.tolist()],
            }

    sp_js_path = os.path.join(BASE, "stacking_plan_data.js")
    with open(sp_js_path, 'w', encoding='utf-8') as f:
        f.write("window.STACKING_PLAN_DATA = ")
        json.dump(sp_out, f, ensure_ascii=False, default=str)
        f.write(";")

    _sp_out_for_enc = sp_out   # capture for GitHub encryption
    sp_size_kb = os.path.getsize(sp_js_path) / 1024
    print(f"✓ stacking_plan_data.js written ({sp_size_kb:.0f} KB), {len(periods)} periods")

except FileNotFoundError as e:
    print(f"  (Stacking plan skipped — file not found: {e})")
except Exception as e:
    print(f"  (Stacking plan error: {e})")

# ══════════════════════════════════════════════════════════════════════
# AUTO-DEPLOY: copy data files to hk-office-360/ and cache-bust index.html
# ══════════════════════════════════════════════════════════════════════
import shutil, re

# If script is already inside hk-office-360 (GitHub Actions), output goes here directly
GITHUB_DIR = BASE if os.path.basename(BASE) == 'hk-office-360' else os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(BASE))), "hk-office-360")
INDEX_HTML  = os.path.join(GITHUB_DIR, "index.html")

if os.path.isdir(GITHUB_DIR):
    # 1. Copy the three data files
    for fname in ("dashboard_data.js", "dashboard_data.json", "stacking_plan_data.js"):
        src = os.path.join(BASE, fname)
        dst = os.path.join(GITHUB_DIR, fname)
        if os.path.exists(src):
            shutil.copy2(src, dst)

    # 2. Update <script src> tags in index.html with ?v=YYYYMMDD (cache-busting)
    if os.path.exists(INDEX_HTML):
        v = datetime.now().strftime("%Y%m%d")
        with open(INDEX_HTML, "r", encoding="utf-8") as f:
            html = f.read()
        html = re.sub(
            r'(<script src="dashboard_data\.js)(?:\?v=\d+)?(")',
            rf'\g<1>?v={v}\g<2>', html)
        html = re.sub(
            r'(<script src="stacking_plan_data\.js)(?:\?v=\d+)?(")',
            rf'\g<1>?v={v}\g<2>', html)
        with open(INDEX_HTML, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"\n✓ hk-office-360/ updated — cache version: v={v}")
    else:
        print(f"\n  (index.html not found in {GITHUB_DIR} — skipped cache-bust)")
else:
    print(f"\n  (hk-office-360/ not found at {GITHUB_DIR} — skipped auto-deploy)")
