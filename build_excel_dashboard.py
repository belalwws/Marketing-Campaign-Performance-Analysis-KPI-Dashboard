import pandas as pd
import numpy as np
import xlsxwriter
import warnings
warnings.filterwarnings('ignore')

base_dir = r"d:\Marketing Campaign Performance Analysis & KPI Dashboard\Marketing-Campaign-Performance-Analysis-KPI-Dashboard"
csv_path = f"{base_dir}\\Marketing campaign dataset.csv"
excel_path = f"{base_dir}\\Campaign_Dashboard.xlsx"

print("Loading and processing data...")
df = pd.read_csv(csv_path)
df = df.drop_duplicates()
df = df[df['impressions'] > 0]
df.loc[df['clicks'] > df['impressions'], 'clicks'] = df['impressions']
df['impressions'] = pd.to_numeric(df['impressions'], errors='coerce')
df['clicks'] = pd.to_numeric(df['clicks'], errors='coerce')
df['media_cost_usd'] = pd.to_numeric(df['media_cost_usd'], errors='coerce')
df['time'] = pd.to_datetime(df['time'])
df['CTR'] = df['clicks'] / df['impressions']
df['CPC'] = df['media_cost_usd'] / df['clicks'].replace(0, np.nan)
df['CPM'] = (df['media_cost_usd'] / df['impressions']) * 1000

target_ctr = 0.02
total_clicks = df['clicks'].sum()
total_impressions = df['impressions'].sum()
total_cost = df['media_cost_usd'].sum()
avg_ctr = total_clicks / total_impressions
avg_cpc = total_cost / total_clicks
avg_cpm = (total_cost / total_impressions) * 1000

# Aggregations
channel_pivot = df.groupby('channel_name').agg(
    Impressions=('impressions', 'sum'), Clicks=('clicks', 'sum'), Cost_USD=('media_cost_usd', 'sum')
).reset_index()
channel_pivot['CTR'] = channel_pivot['Clicks'] / channel_pivot['Impressions']
channel_pivot['CPC'] = channel_pivot['Cost_USD'] / channel_pivot['Clicks']
channel_pivot['CPM'] = (channel_pivot['Cost_USD'] / channel_pivot['Impressions']) * 1000

platform_pivot = df.groupby('ext_service_name').agg(
    Impressions=('impressions', 'sum'), Clicks=('clicks', 'sum'), Cost_USD=('media_cost_usd', 'sum')
).reset_index()
platform_pivot['CTR'] = platform_pivot['Clicks'] / platform_pivot['Impressions']
platform_pivot['CPC'] = platform_pivot['Cost_USD'] / platform_pivot['Clicks']
platform_pivot['CPM'] = (platform_pivot['Cost_USD'] / platform_pivot['Impressions']) * 1000

monthly = df.groupby(df['time'].dt.to_period('M')).agg(
    Impressions=('impressions', 'sum'), Clicks=('clicks', 'sum'), Cost_USD=('media_cost_usd', 'sum')
).reset_index()
monthly['Month'] = monthly['time'].astype(str)
monthly['CTR'] = monthly['Clicks'] / monthly['Impressions']
monthly['CPC'] = monthly['Cost_USD'] / monthly['Clicks']
monthly['Achievement_%'] = (monthly['CTR'] / target_ctr) * 100
monthly['Cumulative_Clicks'] = monthly['Clicks'].cumsum()
monthly['Cumulative_Cost'] = monthly['Cost_USD'].cumsum()
monthly['Cumulative_Impressions'] = monthly['Impressions'].cumsum()
monthly['Cumulative_CTR'] = monthly['Cumulative_Clicks'] / monthly['Cumulative_Impressions']
monthly['CTR_MoM'] = monthly['CTR'].pct_change() * 100
monthly['Clicks_MoM'] = monthly['Clicks'].pct_change() * 100

keyword_pivot = df.groupby('keywords').agg(
    Clicks=('clicks', 'sum'), Impressions=('impressions', 'sum'), Cost_USD=('media_cost_usd', 'sum')
).reset_index()
keyword_pivot['CTR'] = keyword_pivot['Clicks'] / keyword_pivot['Impressions']
keyword_pivot['CPC'] = keyword_pivot['Cost_USD'] / keyword_pivot['Clicks']
top_kw = keyword_pivot.sort_values('Clicks', ascending=False).head(10).reset_index(drop=True)

weekday_pivot = df.groupby('weekday_cat').agg(
    Impressions=('impressions', 'sum'), Clicks=('clicks', 'sum'), Cost_USD=('media_cost_usd', 'sum')
).reset_index()
weekday_pivot['CTR'] = weekday_pivot['Clicks'] / weekday_pivot['Impressions']

# Channel monthly trend
ch_monthly = df.groupby([df['time'].dt.to_period('M'), 'channel_name']).agg(
    Clicks=('clicks', 'sum'), Impressions=('impressions', 'sum'), Cost_USD=('media_cost_usd', 'sum')
).reset_index()
ch_monthly['Month'] = ch_monthly['time'].astype(str)
ch_monthly['CTR'] = ch_monthly['Clicks'] / ch_monthly['Impressions']

# Budget efficiency
channel_eff = channel_pivot[['channel_name', 'Cost_USD', 'Clicks']].copy()
channel_eff['Cost_Share'] = channel_eff['Cost_USD'] / total_cost
channel_eff['Click_Share'] = channel_eff['Clicks'] / total_clicks
channel_eff['Efficiency'] = channel_eff['Click_Share'] / channel_eff['Cost_Share']

# Scorecard
scorecard_ch = channel_pivot[['channel_name', 'CTR']].copy()
scorecard_ch['Target'] = target_ctr
scorecard_ch['Achievement'] = scorecard_ch['CTR'] / target_ctr

scorecard_pl = platform_pivot[['ext_service_name', 'CTR']].copy()
scorecard_pl['Target'] = target_ctr
scorecard_pl['Achievement'] = scorecard_pl['CTR'] / target_ctr

print("Building Excel dashboard...")

wb = xlsxwriter.Workbook(excel_path)

# ===== FORMATS =====
title_fmt = wb.add_format({'bold': True, 'font_size': 20, 'font_color': '#1a1a2e', 'bottom': 2, 'bottom_color': '#2196F3'})
subtitle_fmt = wb.add_format({'bold': True, 'font_size': 14, 'font_color': '#333'})
header_fmt = wb.add_format({'bold': True, 'bg_color': '#2196F3', 'font_color': 'white', 'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'})
header_dark = wb.add_format({'bold': True, 'bg_color': '#1a1a2e', 'font_color': 'white', 'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'})
num_fmt = wb.add_format({'num_format': '#,##0', 'align': 'center', 'border': 1})
pct_fmt = wb.add_format({'num_format': '0.00%', 'align': 'center', 'border': 1})
money_fmt = wb.add_format({'num_format': '$#,##0.00', 'align': 'center', 'border': 1})
text_fmt = wb.add_format({'align': 'center', 'border': 1})
kpi_val_fmt = wb.add_format({'bold': True, 'font_size': 28, 'align': 'center', 'valign': 'vcenter'})
kpi_label_fmt = wb.add_format({'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'font_color': '#666'})
kpi_box_blue = wb.add_format({'bold': True, 'font_size': 28, 'align': 'center', 'valign': 'vcenter', 'font_color': '#2196F3', 'border': 2, 'border_color': '#e0e0e0', 'bg_color': '#f8f9fa'})
kpi_box_green = wb.add_format({'bold': True, 'font_size': 28, 'align': 'center', 'valign': 'vcenter', 'font_color': '#4CAF50', 'border': 2, 'border_color': '#e0e0e0', 'bg_color': '#f8f9fa'})
kpi_box_pink = wb.add_format({'bold': True, 'font_size': 28, 'align': 'center', 'valign': 'vcenter', 'font_color': '#E91E63', 'border': 2, 'border_color': '#e0e0e0', 'bg_color': '#f8f9fa'})
kpi_box_orange = wb.add_format({'bold': True, 'font_size': 28, 'align': 'center', 'valign': 'vcenter', 'font_color': '#FF9800', 'border': 2, 'border_color': '#e0e0e0', 'bg_color': '#f8f9fa'})
kpi_box_purple = wb.add_format({'bold': True, 'font_size': 28, 'align': 'center', 'valign': 'vcenter', 'font_color': '#9C27B0', 'border': 2, 'border_color': '#e0e0e0', 'bg_color': '#f8f9fa'})
kpi_box_cyan = wb.add_format({'bold': True, 'font_size': 28, 'align': 'center', 'valign': 'vcenter', 'font_color': '#00BCD4', 'border': 2, 'border_color': '#e0e0e0', 'bg_color': '#f8f9fa'})
green_fmt = wb.add_format({'bg_color': '#E8F5E9', 'font_color': '#2E7D32', 'bold': True, 'align': 'center', 'border': 1})
red_fmt = wb.add_format({'bg_color': '#FFEBEE', 'font_color': '#C62828', 'bold': True, 'align': 'center', 'border': 1})
pct_green = wb.add_format({'num_format': '0.00%', 'bg_color': '#E8F5E9', 'font_color': '#2E7D32', 'bold': True, 'align': 'center', 'border': 1})
pct_red = wb.add_format({'num_format': '0.00%', 'bg_color': '#FFEBEE', 'font_color': '#C62828', 'bold': True, 'align': 'center', 'border': 1})

channel_colors = ['#2196F3', '#FF9800', '#4CAF50', '#E91E63', '#9C27B0']
platform_colors = ['#1877F2', '#4285F4', '#34A853']

# ============================================================
# SHEET 1: DASHBOARD (KPI Cards + Main Charts)
# ============================================================
ws = wb.add_worksheet('Dashboard')
ws.hide_gridlines(2)
ws.set_tab_color('#2196F3')
ws.set_column('A:A', 3)
ws.set_column('B:M', 16)

ws.merge_range('B1:M1', 'Marketing Campaign KPI Dashboard', title_fmt)
ws.set_row(0, 40)

# KPI Cards Row
ws.merge_range('B3:C4', f'{total_impressions:,.0f}', kpi_box_blue)
ws.merge_range('B5:C5', 'Total Impressions', kpi_label_fmt)
ws.merge_range('D3:E4', f'{total_clicks:,.0f}', kpi_box_green)
ws.merge_range('D5:E5', 'Total Clicks', kpi_label_fmt)
ws.merge_range('F3:G4', f'${total_cost:,.0f}', kpi_box_pink)
ws.merge_range('F5:G5', 'Total Cost (USD)', kpi_label_fmt)
ws.merge_range('H3:I4', f'{avg_ctr:.2%}', kpi_box_orange)
ws.merge_range('H5:I5', 'Overall CTR', kpi_label_fmt)
ws.merge_range('J3:K4', f'${avg_cpc:.2f}', kpi_box_purple)
ws.merge_range('J5:K5', 'Avg CPC', kpi_label_fmt)
ws.merge_range('L3:M4', f'${avg_cpm:.2f}', kpi_box_cyan)
ws.merge_range('L5:M5', 'Avg CPM', kpi_label_fmt)

# ============================================================
# SHEET 2: Channel Performance (data + charts)
# ============================================================
ws2 = wb.add_worksheet('Channel Performance')
ws2.hide_gridlines(2)
ws2.set_tab_color('#4CAF50')
ws2.set_column('A:A', 14)
ws2.set_column('B:G', 16)

ws2.merge_range('A1:G1', 'Performance by Channel', title_fmt)
ws2.set_row(0, 35)

headers = ['Channel', 'Impressions', 'Clicks', 'Cost (USD)', 'CTR', 'CPC', 'CPM']
for i, h in enumerate(headers):
    ws2.write(2, i, h, header_fmt)

for r, (_, row) in enumerate(channel_pivot.iterrows()):
    ws2.write(3+r, 0, row['channel_name'], text_fmt)
    ws2.write(3+r, 1, row['Impressions'], num_fmt)
    ws2.write(3+r, 2, row['Clicks'], num_fmt)
    ws2.write(3+r, 3, row['Cost_USD'], money_fmt)
    ws2.write(3+r, 4, row['CTR'], pct_fmt)
    ws2.write(3+r, 5, row['CPC'], money_fmt)
    ws2.write(3+r, 6, row['CPM'], money_fmt)

# Chart 1: CTR per Channel
chart_ctr = wb.add_chart({'type': 'column'})
chart_ctr.add_series({
    'name': 'CTR',
    'categories': ['Channel Performance', 3, 0, 7, 0],
    'values': ['Channel Performance', 3, 4, 7, 4],
    'fill': {'color': '#2196F3'},
    'data_labels': {'value': True, 'num_format': '0.00%', 'font': {'bold': True, 'size': 10}},
})
chart_ctr.set_title({'name': 'CTR per Channel'})
chart_ctr.set_y_axis({'name': 'CTR', 'num_format': '0.00%'})
chart_ctr.set_x_axis({'name': 'Channel'})
chart_ctr.set_size({'width': 580, 'height': 380})
chart_ctr.set_legend({'none': True})
chart_ctr.set_style(10)
ws2.insert_chart('A10', chart_ctr)

# Chart 2: Clicks per Channel
chart_clicks = wb.add_chart({'type': 'column'})
chart_clicks.add_series({
    'name': 'Clicks',
    'categories': ['Channel Performance', 3, 0, 7, 0],
    'values': ['Channel Performance', 3, 2, 7, 2],
    'fill': {'color': '#4CAF50'},
    'data_labels': {'value': True, 'num_format': '#,##0', 'font': {'bold': True, 'size': 9}},
})
chart_clicks.set_title({'name': 'Total Clicks per Channel'})
chart_clicks.set_y_axis({'name': 'Clicks', 'num_format': '#,##0'})
chart_clicks.set_size({'width': 580, 'height': 380})
chart_clicks.set_legend({'none': True})
chart_clicks.set_style(10)
ws2.insert_chart('I10', chart_clicks)

# Chart 3: Cost per Channel
chart_cost = wb.add_chart({'type': 'column'})
chart_cost.add_series({
    'name': 'Cost (USD)',
    'categories': ['Channel Performance', 3, 0, 7, 0],
    'values': ['Channel Performance', 3, 3, 7, 3],
    'fill': {'color': '#E91E63'},
    'data_labels': {'value': True, 'num_format': '$#,##0', 'font': {'bold': True, 'size': 9}},
})
chart_cost.set_title({'name': 'Total Cost per Channel (USD)'})
chart_cost.set_y_axis({'name': 'Cost (USD)', 'num_format': '$#,##0'})
chart_cost.set_size({'width': 580, 'height': 380})
chart_cost.set_legend({'none': True})
chart_cost.set_style(10)
ws2.insert_chart('A30', chart_cost)

# Chart 4: CPC per Channel
chart_cpc = wb.add_chart({'type': 'column'})
chart_cpc.add_series({
    'name': 'CPC',
    'categories': ['Channel Performance', 3, 0, 7, 0],
    'values': ['Channel Performance', 3, 5, 7, 5],
    'fill': {'color': '#9C27B0'},
    'data_labels': {'value': True, 'num_format': '$0.00', 'font': {'bold': True, 'size': 10}},
})
chart_cpc.set_title({'name': 'CPC per Channel (USD)'})
chart_cpc.set_y_axis({'name': 'CPC (USD)', 'num_format': '$0.00'})
chart_cpc.set_size({'width': 580, 'height': 380})
chart_cpc.set_legend({'none': True})
chart_cpc.set_style(10)
ws2.insert_chart('I30', chart_cpc)

# Chart 5: Impressions Pie
chart_pie = wb.add_chart({'type': 'pie'})
chart_pie.add_series({
    'name': 'Impressions Share',
    'categories': ['Channel Performance', 3, 0, 7, 0],
    'values': ['Channel Performance', 3, 1, 7, 1],
    'points': [{'fill': {'color': c}} for c in channel_colors],
    'data_labels': {'percentage': True, 'category': True, 'font': {'size': 10}},
})
chart_pie.set_title({'name': 'Impressions Share by Channel'})
chart_pie.set_size({'width': 580, 'height': 380})
chart_pie.set_legend({'position': 'bottom'})
chart_pie.set_style(10)
ws2.insert_chart('A50', chart_pie)

# ============================================================
# SHEET 3: Platform Performance
# ============================================================
ws3 = wb.add_worksheet('Platform Performance')
ws3.hide_gridlines(2)
ws3.set_tab_color('#1877F2')
ws3.set_column('A:A', 16)
ws3.set_column('B:G', 16)

ws3.merge_range('A1:G1', 'Performance by Ad Platform', title_fmt)
ws3.set_row(0, 35)

for i, h in enumerate(headers):
    h_p = h.replace('Channel', 'Platform')
    ws3.write(2, i, h_p, header_fmt)

for r, (_, row) in enumerate(platform_pivot.iterrows()):
    ws3.write(3+r, 0, row['ext_service_name'], text_fmt)
    ws3.write(3+r, 1, row['Impressions'], num_fmt)
    ws3.write(3+r, 2, row['Clicks'], num_fmt)
    ws3.write(3+r, 3, row['Cost_USD'], money_fmt)
    ws3.write(3+r, 4, row['CTR'], pct_fmt)
    ws3.write(3+r, 5, row['CPC'], money_fmt)
    ws3.write(3+r, 6, row['CPM'], money_fmt)

# CTR per Platform
ch_p1 = wb.add_chart({'type': 'column'})
ch_p1.add_series({
    'name': 'CTR',
    'categories': ['Platform Performance', 3, 0, 5, 0],
    'values': ['Platform Performance', 3, 4, 5, 4],
    'fill': {'color': '#1877F2'},
    'data_labels': {'value': True, 'num_format': '0.00%', 'font': {'bold': True, 'size': 11}},
})
ch_p1.set_title({'name': 'CTR per Ad Platform'})
ch_p1.set_y_axis({'num_format': '0.00%'})
ch_p1.set_size({'width': 530, 'height': 380})
ch_p1.set_legend({'none': True})
ch_p1.set_style(10)
ws3.insert_chart('A8', ch_p1)

# Clicks per Platform
ch_p2 = wb.add_chart({'type': 'column'})
ch_p2.add_series({
    'name': 'Clicks',
    'categories': ['Platform Performance', 3, 0, 5, 0],
    'values': ['Platform Performance', 3, 2, 5, 2],
    'fill': {'color': '#4285F4'},
    'data_labels': {'value': True, 'num_format': '#,##0', 'font': {'bold': True, 'size': 10}},
})
ch_p2.set_title({'name': 'Total Clicks per Platform'})
ch_p2.set_y_axis({'num_format': '#,##0'})
ch_p2.set_size({'width': 530, 'height': 380})
ch_p2.set_legend({'none': True})
ch_p2.set_style(10)
ws3.insert_chart('I8', ch_p2)

# Platform Pie
ch_p3 = wb.add_chart({'type': 'pie'})
ch_p3.add_series({
    'name': 'Impressions',
    'categories': ['Platform Performance', 3, 0, 5, 0],
    'values': ['Platform Performance', 3, 1, 5, 1],
    'points': [{'fill': {'color': c}} for c in platform_colors],
    'data_labels': {'percentage': True, 'category': True, 'font': {'size': 11}},
})
ch_p3.set_title({'name': 'Impressions Share by Platform'})
ch_p3.set_size({'width': 530, 'height': 380})
ch_p3.set_style(10)
ws3.insert_chart('A28', ch_p3)

# CPC per Platform
ch_p4 = wb.add_chart({'type': 'column'})
ch_p4.add_series({
    'name': 'CPC',
    'categories': ['Platform Performance', 3, 0, 5, 0],
    'values': ['Platform Performance', 3, 5, 5, 5],
    'fill': {'color': '#34A853'},
    'data_labels': {'value': True, 'num_format': '$0.000', 'font': {'bold': True, 'size': 11}},
})
ch_p4.set_title({'name': 'CPC per Platform (USD)'})
ch_p4.set_y_axis({'num_format': '$0.000'})
ch_p4.set_size({'width': 530, 'height': 380})
ch_p4.set_legend({'none': True})
ch_p4.set_style(10)
ws3.insert_chart('I28', ch_p4)

# ============================================================
# SHEET 4: Monthly Trends
# ============================================================
ws4 = wb.add_worksheet('Monthly Trends')
ws4.hide_gridlines(2)
ws4.set_tab_color('#FF9800')
ws4.set_column('A:A', 12)
ws4.set_column('B:K', 16)

ws4.merge_range('A1:K1', 'Monthly Performance Trends', title_fmt)
ws4.set_row(0, 35)

m_headers = ['Month', 'Impressions', 'Clicks', 'Cost (USD)', 'CTR', 'CPC', 'Achievement %',
             'Cum. Clicks', 'Cum. Cost', 'Cum. CTR', 'CTR MoM %', 'Target CTR']
for i, h in enumerate(m_headers):
    ws4.write(2, i, h, header_fmt)

n_months = len(monthly)
for r in range(n_months):
    row = monthly.iloc[r]
    ws4.write(3+r, 0, row['Month'], text_fmt)
    ws4.write(3+r, 1, row['Impressions'], num_fmt)
    ws4.write(3+r, 2, row['Clicks'], num_fmt)
    ws4.write(3+r, 3, row['Cost_USD'], money_fmt)
    ws4.write(3+r, 4, row['CTR'], pct_fmt)
    ws4.write(3+r, 5, row['CPC'], money_fmt)
    ach = row['Achievement_%']
    ws4.write(3+r, 6, ach / 100, pct_green if ach >= 100 else pct_red)
    ws4.write(3+r, 7, row['Cumulative_Clicks'], num_fmt)
    ws4.write(3+r, 8, row['Cumulative_Cost'], money_fmt)
    ws4.write(3+r, 9, row['Cumulative_CTR'], pct_fmt)
    mom = row['CTR_MoM']
    if pd.notna(mom):
        ws4.write(3+r, 10, mom / 100, pct_green if mom >= 0 else pct_red)
    else:
        ws4.write(3+r, 10, '', text_fmt)
    ws4.write(3+r, 11, target_ctr, pct_fmt)

last_data_row = 2 + n_months  # 0-indexed

# Chart: Monthly CTR line
ch_m1 = wb.add_chart({'type': 'line'})
ch_m1.add_series({
    'name': 'Monthly CTR',
    'categories': ['Monthly Trends', 3, 0, last_data_row, 0],
    'values': ['Monthly Trends', 3, 4, last_data_row, 4],
    'line': {'color': '#2196F3', 'width': 3},
    'marker': {'type': 'circle', 'size': 7, 'fill': {'color': '#2196F3'}},
    'data_labels': {'value': True, 'num_format': '0.00%', 'font': {'bold': True, 'size': 9}},
})
ch_m1.add_series({
    'name': 'Target (2%)',
    'categories': ['Monthly Trends', 3, 0, last_data_row, 0],
    'values': ['Monthly Trends', 3, 11, last_data_row, 11],
    'line': {'color': '#F44336', 'width': 2, 'dash_type': 'dash'},
})
ch_m1.set_title({'name': 'Monthly CTR vs Target'})
ch_m1.set_y_axis({'name': 'CTR', 'num_format': '0.00%'})
ch_m1.set_size({'width': 720, 'height': 400})
ch_m1.set_style(10)
ws4.insert_chart('A13', ch_m1)

# Chart: Monthly Clicks bar
ch_m2 = wb.add_chart({'type': 'column'})
ch_m2.add_series({
    'name': 'Clicks',
    'categories': ['Monthly Trends', 3, 0, last_data_row, 0],
    'values': ['Monthly Trends', 3, 2, last_data_row, 2],
    'fill': {'color': '#4CAF50'},
    'data_labels': {'value': True, 'num_format': '#,##0', 'font': {'size': 8}},
})
ch_m2.set_title({'name': 'Monthly Clicks'})
ch_m2.set_y_axis({'num_format': '#,##0'})
ch_m2.set_size({'width': 720, 'height': 400})
ch_m2.set_legend({'none': True})
ch_m2.set_style(10)
ws4.insert_chart('A34', ch_m2)

# Chart: Monthly Cost bar
ch_m3 = wb.add_chart({'type': 'column'})
ch_m3.add_series({
    'name': 'Cost (USD)',
    'categories': ['Monthly Trends', 3, 0, last_data_row, 0],
    'values': ['Monthly Trends', 3, 3, last_data_row, 3],
    'fill': {'color': '#E91E63'},
    'data_labels': {'value': True, 'num_format': '$#,##0', 'font': {'size': 8}},
})
ch_m3.set_title({'name': 'Monthly Cost (USD)'})
ch_m3.set_y_axis({'num_format': '$#,##0'})
ch_m3.set_size({'width': 720, 'height': 400})
ch_m3.set_legend({'none': True})
ch_m3.set_style(10)
ws4.insert_chart('K34', ch_m3)

# Chart: Achievement % bar
ch_m4 = wb.add_chart({'type': 'column'})
ch_m4.add_series({
    'name': 'Achievement %',
    'categories': ['Monthly Trends', 3, 0, last_data_row, 0],
    'values': ['Monthly Trends', 3, 6, last_data_row, 6],
    'data_labels': {'value': True, 'num_format': '0%', 'font': {'bold': True, 'size': 9}},
    'points': [
        {'fill': {'color': '#4CAF50' if monthly.iloc[i]['Achievement_%'] >= 100 else '#F44336'}}
        for i in range(n_months)
    ],
})
ch_m4.set_title({'name': 'Monthly Achievement vs Target (100% = On Target)'})
ch_m4.set_y_axis({'num_format': '0%'})
ch_m4.set_size({'width': 720, 'height': 400})
ch_m4.set_legend({'none': True})
ch_m4.set_style(10)
ws4.insert_chart('K13', ch_m4)

# Chart: Cumulative CTR line
ch_m5 = wb.add_chart({'type': 'line'})
ch_m5.add_series({
    'name': 'Cumulative CTR',
    'categories': ['Monthly Trends', 3, 0, last_data_row, 0],
    'values': ['Monthly Trends', 3, 9, last_data_row, 9],
    'line': {'color': '#FF9800', 'width': 3},
    'marker': {'type': 'circle', 'size': 7, 'fill': {'color': '#FF9800'}},
    'data_labels': {'value': True, 'num_format': '0.00%', 'font': {'size': 9}},
})
ch_m5.add_series({
    'name': 'Target (2%)',
    'categories': ['Monthly Trends', 3, 0, last_data_row, 0],
    'values': ['Monthly Trends', 3, 11, last_data_row, 11],
    'line': {'color': '#F44336', 'width': 2, 'dash_type': 'dash'},
})
ch_m5.set_title({'name': 'Running Cumulative CTR'})
ch_m5.set_y_axis({'num_format': '0.00%'})
ch_m5.set_size({'width': 720, 'height': 400})
ch_m5.set_style(10)
ws4.insert_chart('A55', ch_m5)

# Chart: Cumulative Clicks area
ch_m6 = wb.add_chart({'type': 'area'})
ch_m6.add_series({
    'name': 'Cumulative Clicks',
    'categories': ['Monthly Trends', 3, 0, last_data_row, 0],
    'values': ['Monthly Trends', 3, 7, last_data_row, 7],
    'fill': {'color': '#2196F3', 'transparency': 60},
    'line': {'color': '#2196F3', 'width': 2.5},
    'marker': {'type': 'circle', 'size': 6, 'fill': {'color': '#2196F3'}},
    'data_labels': {'value': True, 'num_format': '#,##0', 'font': {'size': 8}},
})
ch_m6.set_title({'name': 'Cumulative Clicks Over Time'})
ch_m6.set_y_axis({'num_format': '#,##0'})
ch_m6.set_size({'width': 720, 'height': 400})
ch_m6.set_legend({'none': True})
ch_m6.set_style(10)
ws4.insert_chart('K55', ch_m6)

# ============================================================
# SHEET 5: Channel Trends (CTR over time per channel)
# ============================================================
ws5 = wb.add_worksheet('Channel Trends')
ws5.hide_gridlines(2)
ws5.set_tab_color('#9C27B0')
ws5.set_column('A:A', 12)
ws5.set_column('B:F', 14)

ws5.merge_range('A1:F1', 'Channel CTR Trend Over Time', title_fmt)
ws5.set_row(0, 35)

# Write pivot data: Month | Display | Mobile | Search | Social | Video
ch_trend_pivot = ch_monthly.pivot_table(index='Month', columns='channel_name', values='CTR').reset_index()
ch_names_sorted = sorted(channel_pivot['channel_name'].tolist())

trend_headers = ['Month'] + ch_names_sorted
for i, h in enumerate(trend_headers):
    ws5.write(2, i, h, header_fmt)

for r in range(len(ch_trend_pivot)):
    ws5.write(3+r, 0, ch_trend_pivot['Month'].iloc[r], text_fmt)
    for c, ch_name in enumerate(ch_names_sorted):
        ws5.write(3+r, 1+c, ch_trend_pivot[ch_name].iloc[r], pct_fmt)

n_trend = len(ch_trend_pivot)
last_trend_row = 2 + n_trend

ch_t1 = wb.add_chart({'type': 'line'})
for i, ch_name in enumerate(ch_names_sorted):
    ch_t1.add_series({
        'name': ch_name,
        'categories': ['Channel Trends', 3, 0, last_trend_row, 0],
        'values': ['Channel Trends', 3, 1+i, last_trend_row, 1+i],
        'line': {'color': channel_colors[i], 'width': 2.5},
        'marker': {'type': 'circle', 'size': 6, 'fill': {'color': channel_colors[i]}},
    })
# Write target helper column
for r in range(n_trend):
    ws5.write(3+r, len(ch_names_sorted)+1, target_ctr, pct_fmt)
ch_t1.add_series({
    'name': 'Target (2%)',
    'categories': ['Channel Trends', 3, 0, last_trend_row, 0],
    'values': ['Channel Trends', 3, len(ch_names_sorted)+1, last_trend_row, len(ch_names_sorted)+1],
    'line': {'color': '#F44336', 'width': 2, 'dash_type': 'dash'},
})
ch_t1.set_title({'name': 'Channel CTR Trend Over Time'})
ch_t1.set_y_axis({'name': 'CTR', 'num_format': '0.00%'})
ch_t1.set_x_axis({'name': 'Month'})
ch_t1.set_size({'width': 960, 'height': 480})
ch_t1.set_style(10)
ws5.insert_chart('A13', ch_t1)

# Also write platform trend
pl_monthly = df.groupby([df['time'].dt.to_period('M'), 'ext_service_name']).agg(
    Clicks=('clicks', 'sum'), Impressions=('impressions', 'sum')
).reset_index()
pl_monthly['Month'] = pl_monthly['time'].astype(str)
pl_monthly['CTR'] = pl_monthly['Clicks'] / pl_monthly['Impressions']
pl_trend_pivot = pl_monthly.pivot_table(index='Month', columns='ext_service_name', values='CTR').reset_index()
pl_names = sorted(platform_pivot['ext_service_name'].tolist())

start_row_pl = last_trend_row + 30
ws5.merge_range(start_row_pl, 0, start_row_pl, 3, 'Platform CTR Trend Over Time', subtitle_fmt)

for i, h in enumerate(['Month'] + pl_names):
    ws5.write(start_row_pl+1, i, h, header_fmt)

for r in range(len(pl_trend_pivot)):
    ws5.write(start_row_pl+2+r, 0, pl_trend_pivot['Month'].iloc[r], text_fmt)
    for c, pl_name in enumerate(pl_names):
        ws5.write(start_row_pl+2+r, 1+c, pl_trend_pivot[pl_name].iloc[r], pct_fmt)

n_pl = len(pl_trend_pivot)
ch_t2 = wb.add_chart({'type': 'line'})
for i, pl_name in enumerate(pl_names):
    ch_t2.add_series({
        'name': pl_name,
        'categories': ['Channel Trends', start_row_pl+2, 0, start_row_pl+1+n_pl, 0],
        'values': ['Channel Trends', start_row_pl+2, 1+i, start_row_pl+1+n_pl, 1+i],
        'line': {'color': platform_colors[i], 'width': 2.5},
        'marker': {'type': 'square', 'size': 6, 'fill': {'color': platform_colors[i]}},
    })
# Write platform target helper column
for r in range(n_pl):
    ws5.write(start_row_pl+2+r, len(pl_names)+1, target_ctr, pct_fmt)
ch_t2.add_series({
    'name': 'Target (2%)',
    'categories': ['Channel Trends', start_row_pl+2, 0, start_row_pl+1+n_pl, 0],
    'values': ['Channel Trends', start_row_pl+2, len(pl_names)+1, start_row_pl+1+n_pl, len(pl_names)+1],
    'line': {'color': '#F44336', 'width': 2, 'dash_type': 'dash'},
})
ch_t2.set_title({'name': 'Platform CTR Trend Over Time'})
ch_t2.set_y_axis({'num_format': '0.00%'})
ch_t2.set_size({'width': 960, 'height': 480})
ch_t2.set_style(10)
ws5.insert_chart(start_row_pl+2+n_pl+1, 0, ch_t2)

# ============================================================
# SHEET 6: Top Keywords
# ============================================================
ws6 = wb.add_worksheet('Top Keywords')
ws6.hide_gridlines(2)
ws6.set_tab_color('#E91E63')
ws6.set_column('A:A', 28)
ws6.set_column('B:F', 16)

ws6.merge_range('A1:F1', 'Top 10 Keywords by Clicks', title_fmt)
ws6.set_row(0, 35)

kw_headers = ['Keyword', 'Clicks', 'Impressions', 'Cost (USD)', 'CTR', 'CPC']
for i, h in enumerate(kw_headers):
    ws6.write(2, i, h, header_fmt)

for r in range(len(top_kw)):
    row = top_kw.iloc[r]
    ws6.write(3+r, 0, row['keywords'], text_fmt)
    ws6.write(3+r, 1, row['Clicks'], num_fmt)
    ws6.write(3+r, 2, row['Impressions'], num_fmt)
    ws6.write(3+r, 3, row['Cost_USD'], money_fmt)
    ws6.write(3+r, 4, row['CTR'], pct_fmt)
    ws6.write(3+r, 5, row['CPC'], money_fmt)

last_kw = 2 + len(top_kw)

ch_kw1 = wb.add_chart({'type': 'bar'})
ch_kw1.add_series({
    'name': 'Clicks',
    'categories': ['Top Keywords', 3, 0, last_kw, 0],
    'values': ['Top Keywords', 3, 1, last_kw, 1],
    'fill': {'color': '#4CAF50'},
    'data_labels': {'value': True, 'num_format': '#,##0', 'font': {'bold': True, 'size': 9}},
})
ch_kw1.set_title({'name': 'Top 10 Keywords by Clicks'})
ch_kw1.set_x_axis({'reverse': True})
ch_kw1.set_size({'width': 720, 'height': 420})
ch_kw1.set_legend({'none': True})
ch_kw1.set_style(10)
ws6.insert_chart('A14', ch_kw1)

ch_kw2 = wb.add_chart({'type': 'bar'})
ch_kw2.add_series({
    'name': 'CTR',
    'categories': ['Top Keywords', 3, 0, last_kw, 0],
    'values': ['Top Keywords', 3, 4, last_kw, 4],
    'fill': {'color': '#FF9800'},
    'data_labels': {'value': True, 'num_format': '0.00%', 'font': {'bold': True, 'size': 9}},
})
ch_kw2.set_title({'name': 'Top 10 Keywords by CTR'})
ch_kw2.set_x_axis({'reverse': True})
ch_kw2.set_y_axis({'num_format': '0.00%'})
ch_kw2.set_size({'width': 720, 'height': 420})
ch_kw2.set_legend({'none': True})
ch_kw2.set_style(10)
ws6.insert_chart('K14', ch_kw2)

# ============================================================
# SHEET 7: Weekday vs Weekend
# ============================================================
ws7 = wb.add_worksheet('Weekday vs Weekend')
ws7.hide_gridlines(2)
ws7.set_tab_color('#00BCD4')
ws7.set_column('A:A', 14)
ws7.set_column('B:E', 16)

ws7.merge_range('A1:E1', 'Weekday vs Weekend Performance', title_fmt)
ws7.set_row(0, 35)

wk_headers = ['Day Type', 'Impressions', 'Clicks', 'Cost (USD)', 'CTR']
for i, h in enumerate(wk_headers):
    ws7.write(2, i, h, header_fmt)

for r in range(len(weekday_pivot)):
    row = weekday_pivot.iloc[r]
    ws7.write(3+r, 0, row['weekday_cat'], text_fmt)
    ws7.write(3+r, 1, row['Impressions'], num_fmt)
    ws7.write(3+r, 2, row['Clicks'], num_fmt)
    ws7.write(3+r, 3, row['Cost_USD'], money_fmt)
    ws7.write(3+r, 4, row['CTR'], pct_fmt)

ch_wk1 = wb.add_chart({'type': 'column'})
ch_wk1.add_series({
    'name': 'CTR',
    'categories': ['Weekday vs Weekend', 3, 0, 4, 0],
    'values': ['Weekday vs Weekend', 3, 4, 4, 4],
    'points': [{'fill': {'color': '#FF9800'}}, {'fill': {'color': '#2196F3'}}],
    'data_labels': {'value': True, 'num_format': '0.00%', 'font': {'bold': True, 'size': 12}},
})
ch_wk1.set_title({'name': 'CTR: Weekday vs Weekend'})
ch_wk1.set_y_axis({'num_format': '0.00%'})
ch_wk1.set_size({'width': 480, 'height': 380})
ch_wk1.set_legend({'none': True})
ch_wk1.set_style(10)
ws7.insert_chart('A7', ch_wk1)

ch_wk2 = wb.add_chart({'type': 'column'})
ch_wk2.add_series({
    'name': 'Clicks',
    'categories': ['Weekday vs Weekend', 3, 0, 4, 0],
    'values': ['Weekday vs Weekend', 3, 2, 4, 2],
    'points': [{'fill': {'color': '#FF9800'}}, {'fill': {'color': '#2196F3'}}],
    'data_labels': {'value': True, 'num_format': '#,##0', 'font': {'bold': True, 'size': 11}},
})
ch_wk2.set_title({'name': 'Total Clicks: Weekday vs Weekend'})
ch_wk2.set_y_axis({'num_format': '#,##0'})
ch_wk2.set_size({'width': 480, 'height': 380})
ch_wk2.set_legend({'none': True})
ch_wk2.set_style(10)
ws7.insert_chart('H7', ch_wk2)

# ============================================================
# SHEET 8: Scorecard
# ============================================================
ws8 = wb.add_worksheet('Scorecard')
ws8.hide_gridlines(2)
ws8.set_tab_color('#F44336')
ws8.set_column('A:A', 16)
ws8.set_column('B:D', 16)

ws8.merge_range('A1:D1', 'Performance Scorecard vs Target (CTR = 2%)', title_fmt)
ws8.set_row(0, 35)

# Channel scorecard
ws8.write(3, 0, 'Channel', header_dark)
ws8.write(3, 1, 'Actual CTR', header_dark)
ws8.write(3, 2, 'Target CTR', header_dark)
ws8.write(3, 3, 'Achievement', header_dark)

for r in range(len(scorecard_ch)):
    row = scorecard_ch.iloc[r]
    ws8.write(4+r, 0, row['channel_name'], text_fmt)
    ws8.write(4+r, 1, row['CTR'], pct_fmt)
    ws8.write(4+r, 2, row['Target'], pct_fmt)
    ach = row['Achievement']
    fmt = pct_green if ach >= 1 else pct_red
    ws8.write(4+r, 3, ach, fmt)

# Platform scorecard
ws8.merge_range('F1:I1', 'Platform Scorecard', subtitle_fmt)
ws8.write(3, 5, 'Platform', header_dark)
ws8.write(3, 6, 'Actual CTR', header_dark)
ws8.write(3, 7, 'Target CTR', header_dark)
ws8.write(3, 8, 'Achievement', header_dark)
ws8.set_column('E:E', 3)
ws8.set_column('F:F', 16)
ws8.set_column('G:I', 16)

for r in range(len(scorecard_pl)):
    row = scorecard_pl.iloc[r]
    ws8.write(4+r, 5, row['ext_service_name'], text_fmt)
    ws8.write(4+r, 6, row['CTR'], pct_fmt)
    ws8.write(4+r, 7, row['Target'], pct_fmt)
    ach = row['Achievement']
    fmt = pct_green if ach >= 1 else pct_red
    ws8.write(4+r, 8, ach, fmt)

# Scorecard chart - channel
ch_sc1 = wb.add_chart({'type': 'column'})
ch_sc1.add_series({
    'name': 'Actual CTR',
    'categories': ['Scorecard', 4, 0, 8, 0],
    'values': ['Scorecard', 4, 1, 8, 1],
    'fill': {'color': '#4CAF50'},
    'data_labels': {'value': True, 'num_format': '0.00%', 'font': {'bold': True}},
})
ch_sc1.add_series({
    'name': 'Target CTR',
    'categories': ['Scorecard', 4, 0, 8, 0],
    'values': ['Scorecard', 4, 2, 8, 2],
    'fill': {'color': '#F44336', 'transparency': 50},
})
ch_sc1.set_title({'name': 'Channel: Actual CTR vs Target'})
ch_sc1.set_y_axis({'num_format': '0.00%'})
ch_sc1.set_size({'width': 580, 'height': 380})
ch_sc1.set_style(10)
ws8.insert_chart('A10', ch_sc1)

# Scorecard chart - platform
ch_sc2 = wb.add_chart({'type': 'column'})
ch_sc2.add_series({
    'name': 'Actual CTR',
    'categories': ['Scorecard', 4, 5, 6, 5],
    'values': ['Scorecard', 4, 6, 6, 6],
    'fill': {'color': '#1877F2'},
    'data_labels': {'value': True, 'num_format': '0.00%', 'font': {'bold': True}},
})
ch_sc2.add_series({
    'name': 'Target CTR',
    'categories': ['Scorecard', 4, 5, 6, 5],
    'values': ['Scorecard', 4, 7, 6, 7],
    'fill': {'color': '#F44336', 'transparency': 50},
})
ch_sc2.set_title({'name': 'Platform: Actual CTR vs Target'})
ch_sc2.set_y_axis({'num_format': '0.00%'})
ch_sc2.set_size({'width': 580, 'height': 380})
ch_sc2.set_style(10)
ws8.insert_chart('I10', ch_sc2)

# ============================================================
# SHEET 9: Budget Efficiency
# ============================================================
ws9 = wb.add_worksheet('Budget Efficiency')
ws9.hide_gridlines(2)
ws9.set_tab_color('#673AB7')
ws9.set_column('A:A', 14)
ws9.set_column('B:E', 16)

ws9.merge_range('A1:E1', 'Budget Efficiency Analysis', title_fmt)
ws9.set_row(0, 35)

eff_headers = ['Channel', 'Cost Share', 'Click Share', 'Efficiency Ratio']
for i, h in enumerate(eff_headers):
    ws9.write(2, i, h, header_fmt)

for r in range(len(channel_eff)):
    row = channel_eff.iloc[r]
    ws9.write(3+r, 0, row['channel_name'], text_fmt)
    ws9.write(3+r, 1, row['Cost_Share'], pct_fmt)
    ws9.write(3+r, 2, row['Click_Share'], pct_fmt)
    eff = row['Efficiency']
    fmt = green_fmt if eff >= 1 else red_fmt
    ws9.write(3+r, 3, f'{eff:.3f}x', fmt)

last_eff = 2 + len(channel_eff)

ch_ef1 = wb.add_chart({'type': 'column'})
ch_ef1.add_series({
    'name': 'Cost Share',
    'categories': ['Budget Efficiency', 3, 0, last_eff, 0],
    'values': ['Budget Efficiency', 3, 1, last_eff, 1],
    'fill': {'color': '#F44336'},
    'data_labels': {'value': True, 'num_format': '0.0%', 'font': {'bold': True}},
})
ch_ef1.add_series({
    'name': 'Click Share',
    'categories': ['Budget Efficiency', 3, 0, last_eff, 0],
    'values': ['Budget Efficiency', 3, 2, last_eff, 2],
    'fill': {'color': '#4CAF50'},
    'data_labels': {'value': True, 'num_format': '0.0%', 'font': {'bold': True}},
})
ch_ef1.set_title({'name': 'Cost Share vs Click Share per Channel'})
ch_ef1.set_y_axis({'num_format': '0.0%'})
ch_ef1.set_size({'width': 720, 'height': 420})
ch_ef1.set_style(10)
ws9.insert_chart('A9', ch_ef1)

ws9.write(3 + len(channel_eff) + 2, 0, 'Efficiency Ratio > 1 = getting MORE clicks than budget share (good)', green_fmt)
ws9.write(3 + len(channel_eff) + 3, 0, 'Efficiency Ratio < 1 = getting FEWER clicks than budget share (bad)', red_fmt)

# ============================================================
# SHEET 10: Clean Data (first 5000 rows for reference)
# ============================================================
ws10 = wb.add_worksheet('Clean Data')
ws10.set_tab_color('#607D8B')

data_headers = ['Date', 'Platform', 'Channel', 'Keyword', 'Impressions', 'Clicks', 'Cost (USD)', 'Day Type', 'CTR', 'CPC', 'CPM']
for i, h in enumerate(data_headers):
    ws10.write(0, i, h, header_fmt)

export_df = df[['time', 'ext_service_name', 'channel_name', 'keywords', 'impressions', 'clicks',
                'media_cost_usd', 'weekday_cat', 'CTR', 'CPC', 'CPM']].head(5000)

for r in range(len(export_df)):
    row = export_df.iloc[r]
    ws10.write(1+r, 0, str(row['time'].date()), text_fmt)
    ws10.write(1+r, 1, row['ext_service_name'], text_fmt)
    ws10.write(1+r, 2, row['channel_name'], text_fmt)
    ws10.write(1+r, 3, row['keywords'], text_fmt)
    ws10.write(1+r, 4, row['impressions'], num_fmt)
    ws10.write(1+r, 5, row['clicks'], num_fmt)
    ws10.write(1+r, 6, row['media_cost_usd'], money_fmt)
    ws10.write(1+r, 7, row['weekday_cat'], text_fmt)
    ws10.write(1+r, 8, row['CTR'], pct_fmt)
    ws10.write(1+r, 9, row['CPC'] if pd.notna(row['CPC']) else 0, money_fmt)
    ws10.write(1+r, 10, row['CPM'], money_fmt)

ws10.set_column('A:A', 12)
ws10.set_column('B:C', 14)
ws10.set_column('D:D', 28)
ws10.set_column('E:K', 14)
ws10.autofilter(0, 0, len(export_df), 10)

# ===== Add embedded charts to Dashboard sheet =====
# Now go back and add the main charts to the Dashboard tab

# CTR per Channel on Dashboard
dash_ch1 = wb.add_chart({'type': 'column'})
dash_ch1.add_series({
    'name': 'CTR',
    'categories': ['Channel Performance', 3, 0, 7, 0],
    'values': ['Channel Performance', 3, 4, 7, 4],
    'points': [{'fill': {'color': c}} for c in channel_colors],
    'data_labels': {'value': True, 'num_format': '0.00%', 'font': {'bold': True, 'size': 10}},
})
dash_ch1.set_title({'name': 'CTR per Channel'})
dash_ch1.set_y_axis({'num_format': '0.00%'})
dash_ch1.set_size({'width': 480, 'height': 320})
dash_ch1.set_legend({'none': True})
dash_ch1.set_style(10)
ws.insert_chart('B7', dash_ch1)

# CTR per Platform on Dashboard
dash_ch2 = wb.add_chart({'type': 'column'})
dash_ch2.add_series({
    'name': 'CTR',
    'categories': ['Platform Performance', 3, 0, 5, 0],
    'values': ['Platform Performance', 3, 4, 5, 4],
    'points': [{'fill': {'color': c}} for c in platform_colors],
    'data_labels': {'value': True, 'num_format': '0.00%', 'font': {'bold': True, 'size': 11}},
})
dash_ch2.set_title({'name': 'CTR per Platform'})
dash_ch2.set_y_axis({'num_format': '0.00%'})
dash_ch2.set_size({'width': 480, 'height': 320})
dash_ch2.set_legend({'none': True})
dash_ch2.set_style(10)
ws.insert_chart('H7', dash_ch2)

# Monthly CTR on Dashboard
dash_ch3 = wb.add_chart({'type': 'line'})
dash_ch3.add_series({
    'name': 'CTR',
    'categories': ['Monthly Trends', 3, 0, last_data_row, 0],
    'values': ['Monthly Trends', 3, 4, last_data_row, 4],
    'line': {'color': '#2196F3', 'width': 3},
    'marker': {'type': 'circle', 'size': 6, 'fill': {'color': '#2196F3'}},
    'data_labels': {'value': True, 'num_format': '0.00%', 'font': {'size': 8}},
})
dash_ch3.add_series({
    'name': 'Target',
    'categories': ['Monthly Trends', 3, 0, last_data_row, 0],
    'values': ['Monthly Trends', 3, 11, last_data_row, 11],
    'line': {'color': '#F44336', 'width': 2, 'dash_type': 'dash'},
})
dash_ch3.set_title({'name': 'Monthly CTR Trend'})
dash_ch3.set_y_axis({'num_format': '0.00%'})
dash_ch3.set_size({'width': 960, 'height': 340})
dash_ch3.set_style(10)
ws.insert_chart('B23', dash_ch3)

# Monthly Achievement on Dashboard
dash_ch4 = wb.add_chart({'type': 'column'})
dash_ch4.add_series({
    'name': 'Achievement',
    'categories': ['Monthly Trends', 3, 0, last_data_row, 0],
    'values': ['Monthly Trends', 3, 6, last_data_row, 6],
    'data_labels': {'value': True, 'num_format': '0%', 'font': {'bold': True, 'size': 9}},
    'points': [
        {'fill': {'color': '#4CAF50' if monthly.iloc[i]['Achievement_%'] >= 100 else '#F44336'}}
        for i in range(n_months)
    ],
})
dash_ch4.set_title({'name': 'Monthly Achievement vs Target'})
dash_ch4.set_y_axis({'num_format': '0%'})
dash_ch4.set_size({'width': 960, 'height': 340})
dash_ch4.set_legend({'none': True})
dash_ch4.set_style(10)
ws.insert_chart('B40', dash_ch4)

# Top Keywords on Dashboard
dash_ch5 = wb.add_chart({'type': 'bar'})
dash_ch5.add_series({
    'name': 'Clicks',
    'categories': ['Top Keywords', 3, 0, last_kw, 0],
    'values': ['Top Keywords', 3, 1, last_kw, 1],
    'fill': {'color': '#4CAF50'},
    'data_labels': {'value': True, 'num_format': '#,##0', 'font': {'bold': True, 'size': 9}},
})
dash_ch5.set_title({'name': 'Top 10 Keywords'})
dash_ch5.set_x_axis({'reverse': True})
dash_ch5.set_size({'width': 960, 'height': 380})
dash_ch5.set_legend({'none': True})
dash_ch5.set_style(10)
ws.insert_chart('B57', dash_ch5)

# Budget Efficiency on Dashboard
dash_ch6 = wb.add_chart({'type': 'column'})
dash_ch6.add_series({
    'name': 'Cost Share',
    'categories': ['Budget Efficiency', 3, 0, last_eff, 0],
    'values': ['Budget Efficiency', 3, 1, last_eff, 1],
    'fill': {'color': '#F44336'},
})
dash_ch6.add_series({
    'name': 'Click Share',
    'categories': ['Budget Efficiency', 3, 0, last_eff, 0],
    'values': ['Budget Efficiency', 3, 2, last_eff, 2],
    'fill': {'color': '#4CAF50'},
})
dash_ch6.set_title({'name': 'Budget Efficiency: Cost vs Click Share'})
dash_ch6.set_y_axis({'num_format': '0.0%'})
dash_ch6.set_size({'width': 960, 'height': 340})
dash_ch6.set_style(10)
ws.insert_chart('B77', dash_ch6)

# ===== CLOSE =====
wb.close()
print(f"\n✅ Excel Dashboard saved: {excel_path}")
print(f"   📊 10 sheets with 20+ embedded charts")
print(f"   📋 Sheets: Dashboard, Channel Performance, Platform Performance,")
print(f"      Monthly Trends, Channel Trends, Top Keywords, Weekday vs Weekend,")
print(f"      Scorecard, Budget Efficiency, Clean Data")
