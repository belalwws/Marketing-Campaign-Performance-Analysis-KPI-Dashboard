import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# ============================================================
# STEP 1: DATA UNDERSTANDING
# ============================================================
print("=" * 70)
print("STEP 1: DATA UNDERSTANDING")
print("=" * 70)

df = pd.read_csv(r"d:\Marketing Campaign Performance Analysis & KPI Dashboard\Marketing-Campaign-Performance-Analysis-KPI-Dashboard\Marketing campaign dataset.csv")

print(f"Total Records: {len(df):,}")
print(f"Columns: {len(df.columns)}")
print(f"Date Range: {df['time'].min()} to {df['time'].max()}")
print(f"\nAd Platforms: {df['ext_service_name'].unique().tolist()}")
print(f"Channels: {df['channel_name'].unique().tolist()}")
print(f"Unique Keywords: {df['keywords'].nunique()}")
print(f"Unique Campaigns: {df['campaign_item_id'].nunique()}")

# ============================================================
# STEP 2: DATA CLEANING
# ============================================================
print("\n" + "=" * 70)
print("STEP 2: DATA CLEANING")
print("=" * 70)

original_count = len(df)

# Check nulls
null_counts = df[['impressions', 'clicks', 'media_cost_usd', 'channel_name', 'ext_service_name', 'time']].isnull().sum()
print(f"Null values:\n{null_counts}\n")

# Remove duplicates
df = df.drop_duplicates()
dupes_removed = original_count - len(df)
print(f"Duplicates removed: {dupes_removed}")

# Fix: impressions must be > 0
bad_impressions = (df['impressions'] <= 0).sum()
df = df[df['impressions'] > 0]
print(f"Rows with impressions <= 0 removed: {bad_impressions}")

# Fix: clicks can't exceed impressions
bad_clicks = (df['clicks'] > df['impressions']).sum()
df.loc[df['clicks'] > df['impressions'], 'clicks'] = df['impressions']
print(f"Rows where clicks > impressions (capped): {bad_clicks}")

# Ensure numeric
df['impressions'] = pd.to_numeric(df['impressions'], errors='coerce')
df['clicks'] = pd.to_numeric(df['clicks'], errors='coerce')
df['media_cost_usd'] = pd.to_numeric(df['media_cost_usd'], errors='coerce')
df['time'] = pd.to_datetime(df['time'])

print(f"\nClean dataset: {len(df):,} rows")

# ============================================================
# STEP 3: CREATE KPIs
# ============================================================
print("\n" + "=" * 70)
print("STEP 3: KPI CALCULATIONS")
print("=" * 70)

# CTR = clicks / impressions
df['CTR'] = df['clicks'] / df['impressions']

# CPC = media_cost_usd / clicks (handle division by zero)
df['CPC'] = df['media_cost_usd'] / df['clicks'].replace(0, np.nan)

# CPM = (media_cost_usd / impressions) * 1000
df['CPM'] = (df['media_cost_usd'] / df['impressions']) * 1000

# Summary KPIs
total_clicks = df['clicks'].sum()
total_impressions = df['impressions'].sum()
total_cost = df['media_cost_usd'].sum()
avg_ctr = total_clicks / total_impressions
avg_cpc = total_cost / total_clicks
avg_cpm = (total_cost / total_impressions) * 1000

print(f"Total Impressions: {total_impressions:,.0f}")
print(f"Total Clicks:      {total_clicks:,.0f}")
print(f"Total Cost (USD):  ${total_cost:,.2f}")
print(f"Overall CTR:       {avg_ctr:.4%}")
print(f"Overall CPC:       ${avg_cpc:.4f}")
print(f"Overall CPM:       ${avg_cpm:.4f}")

# ============================================================
# STEP 4: PIVOT TABLES
# ============================================================
print("\n" + "=" * 70)
print("STEP 4: PIVOT TABLES")
print("=" * 70)

# 4.1 Performance per Channel
print("\n--- 4.1: Performance per Channel (channel_name) ---")
channel_pivot = df.groupby('channel_name').agg(
    Impressions=('impressions', 'sum'),
    Clicks=('clicks', 'sum'),
    Cost_USD=('media_cost_usd', 'sum')
).reset_index()
channel_pivot['CTR'] = channel_pivot['Clicks'] / channel_pivot['Impressions']
channel_pivot['CPC'] = channel_pivot['Cost_USD'] / channel_pivot['Clicks']
channel_pivot['CPM'] = (channel_pivot['Cost_USD'] / channel_pivot['Impressions']) * 1000
print(channel_pivot.to_string(index=False))

# 4.2 Performance per Ad Platform
print("\n--- 4.2: Performance per Ad Platform (ext_service_name) ---")
platform_pivot = df.groupby('ext_service_name').agg(
    Impressions=('impressions', 'sum'),
    Clicks=('clicks', 'sum'),
    Cost_USD=('media_cost_usd', 'sum')
).reset_index()
platform_pivot['CTR'] = platform_pivot['Clicks'] / platform_pivot['Impressions']
platform_pivot['CPC'] = platform_pivot['Cost_USD'] / platform_pivot['Clicks']
platform_pivot['CPM'] = (platform_pivot['Cost_USD'] / platform_pivot['Impressions']) * 1000
print(platform_pivot.to_string(index=False))

# 4.3 Performance over Time (Monthly)
print("\n--- 4.3: Performance over Time (Monthly) ---")
df['Month'] = df['time'].dt.to_period('M')
monthly_pivot = df.groupby('Month').agg(
    Impressions=('impressions', 'sum'),
    Clicks=('clicks', 'sum'),
    Cost_USD=('media_cost_usd', 'sum')
).reset_index()
monthly_pivot['CTR'] = monthly_pivot['Clicks'] / monthly_pivot['Impressions']
monthly_pivot['Month'] = monthly_pivot['Month'].astype(str)
print(monthly_pivot.to_string(index=False))

# 4.4 Top 15 Keywords by Clicks
print("\n--- 4.4: Top 15 Keywords by Clicks ---")
keyword_pivot = df.groupby('keywords').agg(
    Clicks=('clicks', 'sum'),
    Impressions=('impressions', 'sum'),
    Cost_USD=('media_cost_usd', 'sum')
).reset_index()
keyword_pivot['CTR'] = keyword_pivot['Clicks'] / keyword_pivot['Impressions']
keyword_pivot['CPC'] = keyword_pivot['Cost_USD'] / keyword_pivot['Clicks']
top_keywords = keyword_pivot.sort_values('Clicks', ascending=False).head(15)
print(top_keywords.to_string(index=False))

# 4.5 Weekday vs Weekend
print("\n--- 4.5: Weekday vs Weekend Performance ---")
weekday_pivot = df.groupby('weekday_cat').agg(
    Impressions=('impressions', 'sum'),
    Clicks=('clicks', 'sum'),
    Cost_USD=('media_cost_usd', 'sum'),
    Records=('clicks', 'count')
).reset_index()
weekday_pivot['CTR'] = weekday_pivot['Clicks'] / weekday_pivot['Impressions']
weekday_pivot['Avg_Clicks_Per_Day'] = weekday_pivot['Clicks'] / weekday_pivot['Records']
print(weekday_pivot.to_string(index=False))

# ============================================================
# STEP 5: SCORECARD
# ============================================================
print("\n" + "=" * 70)
print("STEP 5: SCORECARD (Target CTR = 2%)")
print("=" * 70)

target_ctr = 0.02
scorecard = channel_pivot[['channel_name', 'CTR']].copy()
scorecard.columns = ['Channel', 'Actual_CTR']
scorecard['Target_CTR'] = target_ctr
scorecard['Achievement_%'] = (scorecard['Actual_CTR'] / scorecard['Target_CTR']) * 100
scorecard['Status'] = scorecard['Achievement_%'].apply(
    lambda x: '✅ On Target' if x >= 100 else ('⚠️ Close' if x >= 75 else '❌ Below Target')
)
print(scorecard.to_string(index=False))

# Platform scorecard too
print("\n--- Scorecard by Ad Platform ---")
scorecard_platform = platform_pivot[['ext_service_name', 'CTR']].copy()
scorecard_platform.columns = ['Platform', 'Actual_CTR']
scorecard_platform['Target_CTR'] = target_ctr
scorecard_platform['Achievement_%'] = (scorecard_platform['Actual_CTR'] / scorecard_platform['Target_CTR']) * 100
scorecard_platform['Status'] = scorecard_platform['Achievement_%'].apply(
    lambda x: '✅ On Target' if x >= 100 else ('⚠️ Close' if x >= 75 else '❌ Below Target')
)
print(scorecard_platform.to_string(index=False))

# ============================================================
# STEP 6: DASHBOARD CHARTS
# ============================================================
print("\n" + "=" * 70)
print("STEP 6: CREATING DASHBOARD CHARTS...")
print("=" * 70)

output_dir = r"d:\Marketing Campaign Performance Analysis & KPI Dashboard\Marketing-Campaign-Performance-Analysis-KPI-Dashboard"

# Color palette
colors_channel = ['#2196F3', '#FF9800', '#4CAF50', '#E91E63', '#9C27B0']
colors_platform = ['#1877F2', '#4285F4', '#34A853']

# ---- FIGURE 1: KPI Summary Dashboard ----
fig, axes = plt.subplots(2, 3, figsize=(18, 10))
fig.suptitle('Marketing Campaign KPI Dashboard', fontsize=20, fontweight='bold', y=0.98)
fig.patch.set_facecolor('#f5f5f5')

# KPI Card 1: Total Impressions
ax = axes[0, 0]
ax.set_facecolor('white')
ax.text(0.5, 0.6, f'{total_impressions:,.0f}', ha='center', va='center', fontsize=28, fontweight='bold', color='#2196F3')
ax.text(0.5, 0.25, 'Total Impressions', ha='center', va='center', fontsize=14, color='#666')
ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis('off')
for spine in ax.spines.values(): spine.set_visible(False)

# KPI Card 2: Total Clicks
ax = axes[0, 1]
ax.set_facecolor('white')
ax.text(0.5, 0.6, f'{total_clicks:,.0f}', ha='center', va='center', fontsize=28, fontweight='bold', color='#4CAF50')
ax.text(0.5, 0.25, 'Total Clicks', ha='center', va='center', fontsize=14, color='#666')
ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis('off')

# KPI Card 3: Total Cost
ax = axes[0, 2]
ax.set_facecolor('white')
ax.text(0.5, 0.6, f'${total_cost:,.0f}', ha='center', va='center', fontsize=28, fontweight='bold', color='#E91E63')
ax.text(0.5, 0.25, 'Total Cost (USD)', ha='center', va='center', fontsize=14, color='#666')
ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis('off')

# KPI Card 4: Avg CTR
ax = axes[1, 0]
ax.set_facecolor('white')
ax.text(0.5, 0.6, f'{avg_ctr:.2%}', ha='center', va='center', fontsize=28, fontweight='bold', color='#FF9800')
ax.text(0.5, 0.25, 'Overall CTR', ha='center', va='center', fontsize=14, color='#666')
ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis('off')

# KPI Card 5: Avg CPC
ax = axes[1, 1]
ax.set_facecolor('white')
ax.text(0.5, 0.6, f'${avg_cpc:.2f}', ha='center', va='center', fontsize=28, fontweight='bold', color='#9C27B0')
ax.text(0.5, 0.25, 'Avg CPC (USD)', ha='center', va='center', fontsize=14, color='#666')
ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis('off')

# KPI Card 6: Avg CPM
ax = axes[1, 2]
ax.set_facecolor('white')
ax.text(0.5, 0.6, f'${avg_cpm:.2f}', ha='center', va='center', fontsize=28, fontweight='bold', color='#00BCD4')
ax.text(0.5, 0.25, 'Avg CPM (USD)', ha='center', va='center', fontsize=14, color='#666')
ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis('off')

plt.tight_layout()
plt.savefig(f'{output_dir}\\01_KPI_Cards.png', dpi=150, bbox_inches='tight', facecolor='#f5f5f5')
plt.close()
print("✅ 01_KPI_Cards.png saved")

# ---- FIGURE 2: CTR per Channel ----
fig, ax = plt.subplots(figsize=(10, 6))
bars = ax.bar(channel_pivot['channel_name'], channel_pivot['CTR'] * 100, color=colors_channel, edgecolor='white', linewidth=1.5)
ax.axhline(y=target_ctr * 100, color='red', linestyle='--', linewidth=2, label=f'Target CTR ({target_ctr:.0%})')
for bar, val in zip(bars, channel_pivot['CTR']):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.02, f'{val:.2%}', ha='center', fontweight='bold', fontsize=11)
ax.set_title('CTR per Channel', fontsize=16, fontweight='bold')
ax.set_ylabel('CTR (%)')
ax.legend(fontsize=12)
ax.set_facecolor('#fafafa')
plt.tight_layout()
plt.savefig(f'{output_dir}\\02_CTR_per_Channel.png', dpi=150, bbox_inches='tight')
plt.close()
print("✅ 02_CTR_per_Channel.png saved")

# ---- FIGURE 3: Clicks & Cost per Channel ----
fig, ax1 = plt.subplots(figsize=(10, 6))
x = np.arange(len(channel_pivot))
width = 0.35
bars1 = ax1.bar(x - width/2, channel_pivot['Clicks'], width, label='Clicks', color='#2196F3', alpha=0.85)
ax2 = ax1.twinx()
bars2 = ax2.bar(x + width/2, channel_pivot['Cost_USD'], width, label='Cost (USD)', color='#E91E63', alpha=0.85)
ax1.set_xlabel('Channel')
ax1.set_ylabel('Clicks', color='#2196F3')
ax2.set_ylabel('Cost (USD)', color='#E91E63')
ax1.set_xticks(x)
ax1.set_xticklabels(channel_pivot['channel_name'])
ax1.set_title('Clicks & Cost per Channel', fontsize=16, fontweight='bold')
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left')
plt.tight_layout()
plt.savefig(f'{output_dir}\\03_Clicks_Cost_per_Channel.png', dpi=150, bbox_inches='tight')
plt.close()
print("✅ 03_Clicks_Cost_per_Channel.png saved")

# ---- FIGURE 4: CTR per Ad Platform ----
fig, ax = plt.subplots(figsize=(9, 6))
bars = ax.bar(platform_pivot['ext_service_name'], platform_pivot['CTR'] * 100, color=colors_platform, edgecolor='white', linewidth=1.5)
ax.axhline(y=target_ctr * 100, color='red', linestyle='--', linewidth=2, label=f'Target CTR ({target_ctr:.0%})')
for bar, val in zip(bars, platform_pivot['CTR']):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.02, f'{val:.2%}', ha='center', fontweight='bold', fontsize=12)
ax.set_title('CTR per Ad Platform', fontsize=16, fontweight='bold')
ax.set_ylabel('CTR (%)')
ax.legend(fontsize=12)
plt.tight_layout()
plt.savefig(f'{output_dir}\\04_CTR_per_Platform.png', dpi=150, bbox_inches='tight')
plt.close()
print("✅ 04_CTR_per_Platform.png saved")

# ---- FIGURE 5: Clicks over Time (Daily) ----
daily_data = df.groupby('time').agg(
    Clicks=('clicks', 'sum'),
    Impressions=('impressions', 'sum'),
    Cost=('media_cost_usd', 'sum')
).reset_index()
daily_data['CTR'] = daily_data['Clicks'] / daily_data['Impressions']

fig, ax = plt.subplots(figsize=(14, 6))
ax.plot(daily_data['time'], daily_data['Clicks'], color='#2196F3', linewidth=1.5, alpha=0.8)
ax.fill_between(daily_data['time'], daily_data['Clicks'], alpha=0.15, color='#2196F3')
# Add 7-day moving average
daily_data['Clicks_MA7'] = daily_data['Clicks'].rolling(7).mean()
ax.plot(daily_data['time'], daily_data['Clicks_MA7'], color='#E91E63', linewidth=2.5, label='7-Day Moving Avg')
ax.set_title('Daily Clicks Over Time', fontsize=16, fontweight='bold')
ax.set_ylabel('Clicks')
ax.set_xlabel('Date')
ax.legend(fontsize=12)
ax.xaxis.set_major_locator(matplotlib.dates.MonthLocator())
ax.xaxis.set_major_formatter(matplotlib.dates.DateFormatter('%b %Y'))
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(f'{output_dir}\\05_Clicks_Over_Time.png', dpi=150, bbox_inches='tight')
plt.close()
print("✅ 05_Clicks_Over_Time.png saved")

# ---- FIGURE 6: Cost over Time ----
fig, ax = plt.subplots(figsize=(14, 6))
ax.plot(daily_data['time'], daily_data['Cost'], color='#E91E63', linewidth=1.5, alpha=0.8)
ax.fill_between(daily_data['time'], daily_data['Cost'], alpha=0.15, color='#E91E63')
daily_data['Cost_MA7'] = daily_data['Cost'].rolling(7).mean()
ax.plot(daily_data['time'], daily_data['Cost_MA7'], color='#4CAF50', linewidth=2.5, label='7-Day Moving Avg')
ax.set_title('Daily Cost Over Time (USD)', fontsize=16, fontweight='bold')
ax.set_ylabel('Cost (USD)')
ax.set_xlabel('Date')
ax.legend(fontsize=12)
ax.xaxis.set_major_locator(matplotlib.dates.MonthLocator())
ax.xaxis.set_major_formatter(matplotlib.dates.DateFormatter('%b %Y'))
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(f'{output_dir}\\06_Cost_Over_Time.png', dpi=150, bbox_inches='tight')
plt.close()
print("✅ 06_Cost_Over_Time.png saved")

# ---- FIGURE 7: Top 10 Keywords ----
top10 = keyword_pivot.sort_values('Clicks', ascending=True).tail(10)
fig, ax = plt.subplots(figsize=(10, 7))
bars = ax.barh(top10['keywords'], top10['Clicks'], color='#4CAF50', edgecolor='white')
for bar, val in zip(bars, top10['Clicks']):
    ax.text(bar.get_width() + 50, bar.get_y() + bar.get_height()/2, f'{val:,.0f}', va='center', fontweight='bold')
ax.set_title('Top 10 Keywords by Clicks', fontsize=16, fontweight='bold')
ax.set_xlabel('Total Clicks')
plt.tight_layout()
plt.savefig(f'{output_dir}\\07_Top_Keywords.png', dpi=150, bbox_inches='tight')
plt.close()
print("✅ 07_Top_Keywords.png saved")

# ---- FIGURE 8: Weekday vs Weekend ----
fig, axes = plt.subplots(1, 2, figsize=(12, 5))
wk_colors = ['#FF9800', '#2196F3']

ax = axes[0]
ax.bar(weekday_pivot['weekday_cat'], weekday_pivot['CTR'] * 100, color=wk_colors)
for i, val in enumerate(weekday_pivot['CTR']):
    ax.text(i, val*100 + 0.01, f'{val:.2%}', ha='center', fontweight='bold', fontsize=12)
ax.set_title('CTR: Weekday vs Weekend', fontsize=14, fontweight='bold')
ax.set_ylabel('CTR (%)')

ax = axes[1]
ax.bar(weekday_pivot['weekday_cat'], weekday_pivot['Clicks'], color=wk_colors)
for i, val in enumerate(weekday_pivot['Clicks']):
    ax.text(i, val + 100, f'{val:,.0f}', ha='center', fontweight='bold', fontsize=12)
ax.set_title('Total Clicks: Weekday vs Weekend', fontsize=14, fontweight='bold')
ax.set_ylabel('Clicks')

plt.suptitle('Weekday vs Weekend Performance', fontsize=16, fontweight='bold')
plt.tight_layout()
plt.savefig(f'{output_dir}\\08_Weekday_vs_Weekend.png', dpi=150, bbox_inches='tight')
plt.close()
print("✅ 08_Weekday_vs_Weekend.png saved")

# ---- FIGURE 9: CPC per Channel & Platform ----
fig, axes = plt.subplots(1, 2, figsize=(14, 6))

ax = axes[0]
bars = ax.bar(channel_pivot['channel_name'], channel_pivot['CPC'], color=colors_channel)
for bar, val in zip(bars, channel_pivot['CPC']):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.05, f'${val:.2f}', ha='center', fontweight='bold')
ax.set_title('CPC per Channel', fontsize=14, fontweight='bold')
ax.set_ylabel('CPC (USD)')

ax = axes[1]
bars = ax.bar(platform_pivot['ext_service_name'], platform_pivot['CPC'], color=colors_platform)
for bar, val in zip(bars, platform_pivot['CPC']):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.05, f'${val:.2f}', ha='center', fontweight='bold')
ax.set_title('CPC per Platform', fontsize=14, fontweight='bold')
ax.set_ylabel('CPC (USD)')

plt.suptitle('Cost Per Click Comparison', fontsize=16, fontweight='bold')
plt.tight_layout()
plt.savefig(f'{output_dir}\\09_CPC_Comparison.png', dpi=150, bbox_inches='tight')
plt.close()
print("✅ 09_CPC_Comparison.png saved")

# ---- FIGURE 10: Impressions Share (Pie Charts) ----
fig, axes = plt.subplots(1, 2, figsize=(14, 6))

ax = axes[0]
ax.pie(channel_pivot['Impressions'], labels=channel_pivot['channel_name'], autopct='%1.1f%%',
       colors=colors_channel, startangle=90, textprops={'fontsize': 11})
ax.set_title('Impressions Share by Channel', fontsize=14, fontweight='bold')

ax = axes[1]
ax.pie(platform_pivot['Impressions'], labels=platform_pivot['ext_service_name'], autopct='%1.1f%%',
       colors=colors_platform, startangle=90, textprops={'fontsize': 11})
ax.set_title('Impressions Share by Platform', fontsize=14, fontweight='bold')

plt.tight_layout()
plt.savefig(f'{output_dir}\\10_Impressions_Share.png', dpi=150, bbox_inches='tight')
plt.close()
print("✅ 10_Impressions_Share.png saved")

# ---- FIGURE 11: Efficiency Matrix (CTR vs CPC by keyword) ----
fig, ax = plt.subplots(figsize=(12, 8))
kw_data = keyword_pivot.copy()
kw_data = kw_data[kw_data['CPC'].notna()]
scatter = ax.scatter(kw_data['CPC'], kw_data['CTR']*100, s=kw_data['Clicks']/5, 
                     alpha=0.6, c=kw_data['Clicks'], cmap='viridis', edgecolors='white')
ax.set_xlabel('CPC (USD)', fontsize=12)
ax.set_ylabel('CTR (%)', fontsize=12)
ax.set_title('Keyword Efficiency Matrix (Size = Volume of Clicks)', fontsize=16, fontweight='bold')
# Highlight top 5 keywords
for _, row in top_keywords.head(5).iterrows():
    if pd.notna(row['CPC']):
        ax.annotate(row['keywords'], (row['CPC'], row['CTR']*100), fontsize=8, alpha=0.8)
plt.colorbar(scatter, label='Total Clicks')
ax.axhline(y=target_ctr*100, color='red', linestyle='--', alpha=0.5, label='Target CTR')
ax.legend()
plt.tight_layout()
plt.savefig(f'{output_dir}\\11_Efficiency_Matrix.png', dpi=150, bbox_inches='tight')
plt.close()
print("✅ 11_Efficiency_Matrix.png saved")

# ============================================================
# STEP 7: EXPORT TO EXCEL
# ============================================================
print("\n" + "=" * 70)
print("STEP 7: EXPORTING TO EXCEL")
print("=" * 70)

excel_path = f'{output_dir}\\Campaign_Analysis_Report.xlsx'
with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
    # Clean data sheet
    df_export = df[['time', 'ext_service_name', 'channel_name', 'keywords', 'impressions', 'clicks', 
                    'media_cost_usd', 'weekday_cat', 'CTR', 'CPC', 'CPM']].copy()
    df_export.to_excel(writer, sheet_name='Clean Data', index=False)
    
    # Channel Performance
    channel_pivot.to_excel(writer, sheet_name='Channel Performance', index=False)
    
    # Platform Performance
    platform_pivot.to_excel(writer, sheet_name='Platform Performance', index=False)
    
    # Monthly Performance
    monthly_pivot.to_excel(writer, sheet_name='Monthly Performance', index=False)
    
    # Top Keywords
    keyword_pivot.sort_values('Clicks', ascending=False).to_excel(writer, sheet_name='Keywords', index=False)
    
    # Weekday vs Weekend
    weekday_pivot.to_excel(writer, sheet_name='Weekday vs Weekend', index=False)
    
    # Scorecard
    scorecard.to_excel(writer, sheet_name='Scorecard - Channel', index=False)
    scorecard_platform.to_excel(writer, sheet_name='Scorecard - Platform', index=False)
    
    # Format sheets
    workbook = writer.book
    header_format = workbook.add_format({
        'bold': True, 'bg_color': '#2196F3', 'font_color': 'white',
        'border': 1, 'text_wrap': True, 'valign': 'vcenter'
    })
    pct_format = workbook.add_format({'num_format': '0.00%'})
    money_format = workbook.add_format({'num_format': '$#,##0.00'})
    number_format = workbook.add_format({'num_format': '#,##0'})
    
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        worksheet.set_column('A:Z', 18)

print(f"✅ Excel report saved: {excel_path}")

# ============================================================
# STEP 8: INSIGHTS SUMMARY
# ============================================================
print("\n" + "=" * 70)
print("STEP 8: KEY INSIGHTS FOR INTERVIEW")
print("=" * 70)

# Find best/worst channels
best_ctr_channel = channel_pivot.loc[channel_pivot['CTR'].idxmax()]
worst_ctr_channel = channel_pivot.loc[channel_pivot['CTR'].idxmin()]
cheapest_channel = channel_pivot.loc[channel_pivot['CPC'].idxmin()]
most_expensive_channel = channel_pivot.loc[channel_pivot['CPC'].idxmax()]

best_ctr_platform = platform_pivot.loc[platform_pivot['CTR'].idxmax()]
worst_ctr_platform = platform_pivot.loc[platform_pivot['CTR'].idxmin()]

# Weekday vs Weekend insight
wk_day_ctr = weekday_pivot[weekday_pivot['weekday_cat'] == 'week_day']['CTR'].values[0]
wk_end_ctr = weekday_pivot[weekday_pivot['weekday_cat'] == 'week_end']['CTR'].values[0]

# Most efficient keywords (high CTR, low CPC)
efficient_kw = keyword_pivot[(keyword_pivot['Clicks'] > 100)].copy()
efficient_kw['Efficiency'] = efficient_kw['CTR'] / efficient_kw['CPC'].replace(0, np.nan)
top_efficient = efficient_kw.sort_values('Efficiency', ascending=False).head(5)

print(f"""
📊 INSIGHT 1 - Channel Performance:
   Best CTR Channel:  {best_ctr_channel['channel_name']} ({best_ctr_channel['CTR']:.2%})
   Worst CTR Channel: {worst_ctr_channel['channel_name']} ({worst_ctr_channel['CTR']:.2%})
   Cheapest CPC:      {cheapest_channel['channel_name']} (${cheapest_channel['CPC']:.2f})
   Most Expensive:    {most_expensive_channel['channel_name']} (${most_expensive_channel['CPC']:.2f})

📊 INSIGHT 2 - Platform Performance:
   Best CTR Platform:  {best_ctr_platform['ext_service_name']} ({best_ctr_platform['CTR']:.2%})
   Worst CTR Platform: {worst_ctr_platform['ext_service_name']} ({worst_ctr_platform['CTR']:.2%})

📊 INSIGHT 3 - Timing:
   Weekday CTR: {wk_day_ctr:.2%}
   Weekend CTR: {wk_end_ctr:.2%}
   {'Weekdays perform better' if wk_day_ctr > wk_end_ctr else 'Weekends perform better'}

📊 INSIGHT 4 - Top Efficient Keywords (High CTR, Low CPC):""")

for _, row in top_efficient.iterrows():
    print(f"   • {row['keywords']}: CTR={row['CTR']:.2%}, CPC=${row['CPC']:.2f}, Clicks={row['Clicks']:,.0f}")

print(f"""
📊 INSIGHT 5 - Budget Optimization:
   Total budget spent: ${total_cost:,.2f}
   If we reallocate budget from {worst_ctr_channel['channel_name']} to {best_ctr_channel['channel_name']},
   we could potentially improve overall CTR by {((best_ctr_channel['CTR'] - worst_ctr_channel['CTR']) / worst_ctr_channel['CTR'] * 100):.0f}%

🔥 GOLDEN INTERVIEW STATEMENT:
   "I analyzed {len(df):,} campaign records across {df['ext_service_name'].nunique()} ad platforms
   and {df['channel_name'].nunique()} channels. I identified that {best_ctr_channel['channel_name']} delivers
   the highest engagement (CTR: {best_ctr_channel['CTR']:.2%}), while {cheapest_channel['channel_name']} offers
   the most cost-effective clicks at ${cheapest_channel['CPC']:.2f} CPC. This analysis can help optimize
   budget allocation to maximize ROI across channels."
""")

print("=" * 70)
print("ALL DONE! Files created:")
print("=" * 70)

# ============================================================
# STEP 9: ACHIEVEMENT TREND & TARGET TRACKING (Senior Analyst Layer)
# ============================================================
print("\n" + "=" * 70)
print("STEP 9: ACHIEVEMENT TREND & TARGET TRACKING")
print("=" * 70)

# --- 9.1: Monthly Achievement vs Target ---
monthly_achievement = df.groupby(df['time'].dt.to_period('M')).agg(
    Impressions=('impressions', 'sum'),
    Clicks=('clicks', 'sum'),
    Cost_USD=('media_cost_usd', 'sum')
).reset_index()
monthly_achievement['Month'] = monthly_achievement['time'].astype(str)
monthly_achievement['CTR'] = monthly_achievement['Clicks'] / monthly_achievement['Impressions']
monthly_achievement['CPC'] = monthly_achievement['Cost_USD'] / monthly_achievement['Clicks']
monthly_achievement['Achievement_%'] = (monthly_achievement['CTR'] / target_ctr) * 100
monthly_achievement['Status'] = monthly_achievement['Achievement_%'].apply(
    lambda x: 'Above Target' if x >= 100 else 'Below Target'
)

# Cumulative metrics (running totals)
monthly_achievement['Cumulative_Clicks'] = monthly_achievement['Clicks'].cumsum()
monthly_achievement['Cumulative_Cost'] = monthly_achievement['Cost_USD'].cumsum()
monthly_achievement['Cumulative_Impressions'] = monthly_achievement['Impressions'].cumsum()
monthly_achievement['Cumulative_CTR'] = monthly_achievement['Cumulative_Clicks'] / monthly_achievement['Cumulative_Impressions']
monthly_achievement['Cumulative_Achievement_%'] = (monthly_achievement['Cumulative_CTR'] / target_ctr) * 100

# MoM Growth
monthly_achievement['CTR_MoM_Change'] = monthly_achievement['CTR'].pct_change() * 100
monthly_achievement['Clicks_MoM_Change'] = monthly_achievement['Clicks'].pct_change() * 100
monthly_achievement['Cost_MoM_Change'] = monthly_achievement['Cost_USD'].pct_change() * 100

print("\n--- Monthly Achievement Tracking ---")
display_cols = ['Month', 'Clicks', 'Cost_USD', 'CTR', 'Achievement_%', 'Status', 'CTR_MoM_Change']
print(monthly_achievement[display_cols].to_string(index=False))

# --- 9.2: Channel × Month Breakdown (Where is the performance coming from?) ---
channel_monthly = df.groupby([df['time'].dt.to_period('M'), 'channel_name']).agg(
    Clicks=('clicks', 'sum'),
    Impressions=('impressions', 'sum'),
    Cost_USD=('media_cost_usd', 'sum')
).reset_index()
channel_monthly['Month'] = channel_monthly['time'].astype(str)
channel_monthly['CTR'] = channel_monthly['Clicks'] / channel_monthly['Impressions']
channel_monthly['Achievement_%'] = (channel_monthly['CTR'] / target_ctr) * 100

print("\n--- Channel Performance Trend (CTR by Month) ---")
channel_trend_display = channel_monthly.pivot_table(index='Month', columns='channel_name', values='CTR')
print((channel_trend_display * 100).round(2).to_string())

# --- 9.3: Platform × Month Breakdown ---
platform_monthly = df.groupby([df['time'].dt.to_period('M'), 'ext_service_name']).agg(
    Clicks=('clicks', 'sum'),
    Impressions=('impressions', 'sum'),
    Cost_USD=('media_cost_usd', 'sum')
).reset_index()
platform_monthly['Month'] = platform_monthly['time'].astype(str)
platform_monthly['CTR'] = platform_monthly['Clicks'] / platform_monthly['Impressions']

print("\n--- Platform Performance Trend (CTR by Month) ---")
platform_trend_display = platform_monthly.pivot_table(index='Month', columns='ext_service_name', values='CTR')
print((platform_trend_display * 100).round(2).to_string())

# --- 9.4: Cost Contribution (Where is the money going?) ---
channel_cost_share = channel_pivot[['channel_name', 'Cost_USD', 'Clicks', 'CTR']].copy()
channel_cost_share['Cost_Share_%'] = (channel_cost_share['Cost_USD'] / total_cost) * 100
channel_cost_share['Click_Share_%'] = (channel_cost_share['Clicks'] / total_clicks) * 100
channel_cost_share['Efficiency_Ratio'] = channel_cost_share['Click_Share_%'] / channel_cost_share['Cost_Share_%']

print("\n--- Cost vs Click Contribution (Efficiency Ratio) ---")
print("   (Efficiency Ratio > 1 = getting MORE clicks than the share of budget spent)")
print("   (Efficiency Ratio < 1 = getting FEWER clicks for the money)")
print(channel_cost_share[['channel_name', 'Cost_Share_%', 'Click_Share_%', 'Efficiency_Ratio']].to_string(index=False))

# ============================================================
# STEP 9 CHARTS
# ============================================================

# ---- FIGURE 12: Monthly Achievement vs Target ----
fig, ax1 = plt.subplots(figsize=(14, 7))

months = monthly_achievement['Month']
x = np.arange(len(months))
achievement = monthly_achievement['Achievement_%']
colors_ach = ['#4CAF50' if a >= 100 else '#F44336' for a in achievement]

bars = ax1.bar(x, achievement, color=colors_ach, edgecolor='white', linewidth=1.5, width=0.6, alpha=0.85, label='Monthly Achievement')
ax1.axhline(y=100, color='#333', linestyle='--', linewidth=2, label='Target (100%)')

for bar, val, ctr in zip(bars, achievement, monthly_achievement['CTR']):
    ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 3,
             f'{val:.0f}%\n({ctr:.2%})', ha='center', fontsize=9, fontweight='bold')

# Cumulative achievement line on secondary axis
ax2 = ax1.twinx()
ax2.plot(x, monthly_achievement['Cumulative_Achievement_%'], color='#FF9800',
         linewidth=3, marker='o', markersize=8, label='Cumulative Achievement', zorder=5)
ax2.set_ylabel('Cumulative Achievement (%)', color='#FF9800', fontsize=12)

ax1.set_xticks(x)
ax1.set_xticklabels(months, rotation=45)
ax1.set_ylabel('Monthly Achievement (%)', fontsize=12)
ax1.set_title('Monthly CTR Achievement vs Target (2%)', fontsize=16, fontweight='bold')
ax1.set_ylim(0, max(achievement) * 1.25)

lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper right', fontsize=11)

plt.tight_layout()
plt.savefig(f'{output_dir}\\12_Achievement_vs_Target.png', dpi=150, bbox_inches='tight')
plt.close()
print("✅ 12_Achievement_vs_Target.png saved")

# ---- FIGURE 13: Channel CTR Trend Over Time (Heatmap-style) ----
fig, ax = plt.subplots(figsize=(14, 7))
channel_names = channel_trend_display.columns.tolist()
for i, ch in enumerate(channel_names):
    vals = channel_trend_display[ch].values * 100
    ax.plot(channel_trend_display.index.astype(str), vals, marker='o', linewidth=2.5,
            markersize=8, label=ch, color=colors_channel[i])
    for j, v in enumerate(vals):
        ax.annotate(f'{v:.1f}%', (j, v), textcoords="offset points",
                    xytext=(0, 10), ha='center', fontsize=8)

ax.axhline(y=target_ctr*100, color='red', linestyle='--', linewidth=2, alpha=0.7, label='Target (2%)')
ax.set_title('Channel CTR Trend Over Time', fontsize=16, fontweight='bold')
ax.set_ylabel('CTR (%)', fontsize=12)
ax.set_xlabel('Month', fontsize=12)
ax.legend(fontsize=10, loc='upper right')
ax.set_facecolor('#fafafa')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(f'{output_dir}\\13_Channel_CTR_Trend.png', dpi=150, bbox_inches='tight')
plt.close()
print("✅ 13_Channel_CTR_Trend.png saved")

# ---- FIGURE 14: Platform CTR Trend Over Time ----
fig, ax = plt.subplots(figsize=(14, 7))
platform_names = platform_trend_display.columns.tolist()
for i, pl in enumerate(platform_names):
    vals = platform_trend_display[pl].values * 100
    ax.plot(platform_trend_display.index.astype(str), vals, marker='s', linewidth=2.5,
            markersize=8, label=pl, color=colors_platform[i])
    for j, v in enumerate(vals):
        ax.annotate(f'{v:.1f}%', (j, v), textcoords="offset points",
                    xytext=(0, 10), ha='center', fontsize=8)

ax.axhline(y=target_ctr*100, color='red', linestyle='--', linewidth=2, alpha=0.7, label='Target (2%)')
ax.set_title('Platform CTR Trend Over Time', fontsize=16, fontweight='bold')
ax.set_ylabel('CTR (%)', fontsize=12)
ax.set_xlabel('Month', fontsize=12)
ax.legend(fontsize=10, loc='upper right')
ax.set_facecolor('#fafafa')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(f'{output_dir}\\14_Platform_CTR_Trend.png', dpi=150, bbox_inches='tight')
plt.close()
print("✅ 14_Platform_CTR_Trend.png saved")

# ---- FIGURE 15: MoM Growth Waterfall ----
fig, axes = plt.subplots(1, 3, figsize=(18, 6))
fig.suptitle('Month-over-Month Growth Rate (%)', fontsize=16, fontweight='bold')

mom_data = monthly_achievement.iloc[1:]  # skip first month (no previous)

# CTR MoM
ax = axes[0]
colors_mom = ['#4CAF50' if v >= 0 else '#F44336' for v in mom_data['CTR_MoM_Change']]
bars = ax.bar(mom_data['Month'], mom_data['CTR_MoM_Change'], color=colors_mom, edgecolor='white')
for bar, val in zip(bars, mom_data['CTR_MoM_Change']):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + (1 if val >= 0 else -3),
            f'{val:+.0f}%', ha='center', fontsize=9, fontweight='bold')
ax.axhline(y=0, color='black', linewidth=1)
ax.set_title('CTR Growth', fontsize=13, fontweight='bold')
ax.set_ylabel('MoM Change (%)')
plt.setp(ax.get_xticklabels(), rotation=45)

# Clicks MoM
ax = axes[1]
colors_mom = ['#4CAF50' if v >= 0 else '#F44336' for v in mom_data['Clicks_MoM_Change']]
bars = ax.bar(mom_data['Month'], mom_data['Clicks_MoM_Change'], color=colors_mom, edgecolor='white')
for bar, val in zip(bars, mom_data['Clicks_MoM_Change']):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + (5 if val >= 0 else -15),
            f'{val:+.0f}%', ha='center', fontsize=9, fontweight='bold')
ax.axhline(y=0, color='black', linewidth=1)
ax.set_title('Clicks Growth', fontsize=13, fontweight='bold')
ax.set_ylabel('MoM Change (%)')
plt.setp(ax.get_xticklabels(), rotation=45)

# Cost MoM
ax = axes[2]
colors_mom = ['#4CAF50' if v >= 0 else '#F44336' for v in mom_data['Cost_MoM_Change']]
bars = ax.bar(mom_data['Month'], mom_data['Cost_MoM_Change'], color=colors_mom, edgecolor='white')
for bar, val in zip(bars, mom_data['Cost_MoM_Change']):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + (5 if val >= 0 else -15),
            f'{val:+.0f}%', ha='center', fontsize=9, fontweight='bold')
ax.axhline(y=0, color='black', linewidth=1)
ax.set_title('Cost Growth', fontsize=13, fontweight='bold')
ax.set_ylabel('MoM Change (%)')
plt.setp(ax.get_xticklabels(), rotation=45)

plt.tight_layout()
plt.savefig(f'{output_dir}\\15_MoM_Growth.png', dpi=150, bbox_inches='tight')
plt.close()
print("✅ 15_MoM_Growth.png saved")

# ---- FIGURE 16: Cost vs Click Share (Efficiency) ----
fig, ax = plt.subplots(figsize=(10, 7))
x_pos = np.arange(len(channel_cost_share))
width = 0.35

bars1 = ax.bar(x_pos - width/2, channel_cost_share['Cost_Share_%'], width,
               label='Cost Share %', color='#F44336', alpha=0.85)
bars2 = ax.bar(x_pos + width/2, channel_cost_share['Click_Share_%'], width,
               label='Click Share %', color='#4CAF50', alpha=0.85)

for bar, val in zip(bars1, channel_cost_share['Cost_Share_%']):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.2,
            f'{val:.1f}%', ha='center', fontweight='bold', fontsize=10)
for bar, val in zip(bars2, channel_cost_share['Click_Share_%']):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.2,
            f'{val:.1f}%', ha='center', fontweight='bold', fontsize=10)

# Add efficiency ratio annotation
for i, (_, row) in enumerate(channel_cost_share.iterrows()):
    ratio = row['Efficiency_Ratio']
    color = '#4CAF50' if ratio >= 1 else '#F44336'
    ax.text(i, max(row['Cost_Share_%'], row['Click_Share_%']) + 1.5,
            f'Efficiency: {ratio:.2f}x', ha='center', fontsize=10,
            fontweight='bold', color=color,
            bbox=dict(boxstyle='round,pad=0.3', facecolor='white', edgecolor=color, alpha=0.9))

ax.set_xticks(x_pos)
ax.set_xticklabels(channel_cost_share['channel_name'])
ax.set_ylabel('Share (%)', fontsize=12)
ax.set_title('Budget Efficiency: Cost Share vs Click Share per Channel', fontsize=16, fontweight='bold')
ax.legend(fontsize=12)
ax.set_facecolor('#fafafa')
plt.tight_layout()
plt.savefig(f'{output_dir}\\16_Budget_Efficiency.png', dpi=150, bbox_inches='tight')
plt.close()
print("✅ 16_Budget_Efficiency.png saved")

# ---- FIGURE 17: Cumulative Performance Over Time ----
fig, axes = plt.subplots(1, 3, figsize=(18, 6))
fig.suptitle('Cumulative Performance Over Time', fontsize=16, fontweight='bold')

ax = axes[0]
ax.fill_between(monthly_achievement['Month'], monthly_achievement['Cumulative_Clicks'],
                alpha=0.3, color='#2196F3')
ax.plot(monthly_achievement['Month'], monthly_achievement['Cumulative_Clicks'],
        marker='o', linewidth=2.5, color='#2196F3')
for i, val in enumerate(monthly_achievement['Cumulative_Clicks']):
    ax.text(i, val + total_clicks*0.02, f'{val/1e6:.1f}M', ha='center', fontsize=9, fontweight='bold')
ax.set_title('Cumulative Clicks', fontsize=13, fontweight='bold')
ax.set_ylabel('Clicks')
plt.setp(ax.get_xticklabels(), rotation=45)

ax = axes[1]
ax.fill_between(monthly_achievement['Month'], monthly_achievement['Cumulative_Cost'],
                alpha=0.3, color='#E91E63')
ax.plot(monthly_achievement['Month'], monthly_achievement['Cumulative_Cost'],
        marker='o', linewidth=2.5, color='#E91E63')
for i, val in enumerate(monthly_achievement['Cumulative_Cost']):
    ax.text(i, val + total_cost*0.02, f'${val/1e3:.0f}K', ha='center', fontsize=9, fontweight='bold')
ax.set_title('Cumulative Cost (USD)', fontsize=13, fontweight='bold')
ax.set_ylabel('Cost (USD)')
plt.setp(ax.get_xticklabels(), rotation=45)

ax = axes[2]
ax.plot(monthly_achievement['Month'], monthly_achievement['Cumulative_CTR'] * 100,
        marker='o', linewidth=2.5, color='#FF9800')
ax.axhline(y=target_ctr*100, color='red', linestyle='--', linewidth=2, label='Target (2%)')
for i, val in enumerate(monthly_achievement['Cumulative_CTR']):
    ax.text(i, val*100 + 0.05, f'{val:.2%}', ha='center', fontsize=9, fontweight='bold')
ax.set_title('Running CTR (Cumulative)', fontsize=13, fontweight='bold')
ax.set_ylabel('CTR (%)')
ax.legend()
plt.setp(ax.get_xticklabels(), rotation=45)

plt.tight_layout()
plt.savefig(f'{output_dir}\\17_Cumulative_Performance.png', dpi=150, bbox_inches='tight')
plt.close()
print("✅ 17_Cumulative_Performance.png saved")

# --- Print Summary ---
print("\n--- Achievement Trend Summary ---")
above_target_months = (monthly_achievement['Achievement_%'] >= 100).sum()
below_target_months = (monthly_achievement['Achievement_%'] < 100).sum()
best_month = monthly_achievement.loc[monthly_achievement['CTR'].idxmax()]
worst_month = monthly_achievement.loc[monthly_achievement['CTR'].idxmin()]

print(f"   Months ABOVE target: {above_target_months}/8")
print(f"   Months BELOW target: {below_target_months}/8")
print(f"   Best Month:  {best_month['Month']} (CTR={best_month['CTR']:.2%}, Achievement={best_month['Achievement_%']:.0f}%)")
print(f"   Worst Month: {worst_month['Month']} (CTR={worst_month['CTR']:.2%}, Achievement={worst_month['Achievement_%']:.0f}%)")
print(f"   Cumulative CTR: {monthly_achievement['Cumulative_CTR'].iloc[-1]:.2%}")
print(f"   Overall Achievement: {monthly_achievement['Cumulative_Achievement_%'].iloc[-1]:.0f}%")

# Which channel CONSISTENTLY beats target?
print("\n--- Channel Consistency (months above target) ---")
for ch in channel_names:
    ch_data = channel_monthly[channel_monthly['channel_name'] == ch]
    above = (ch_data['Achievement_%'] >= 100).sum()
    print(f"   {ch}: {above}/8 months above target")

# --- Update Excel with new sheets ---
print("\n--- Updating Excel with Achievement Tracking sheets ---")
import openpyxl
wb = openpyxl.load_workbook(excel_path)

# Remove sheets if they exist (re-run safe)
for sheet_name in ['Monthly Achievement', 'Channel Trend', 'Budget Efficiency']:
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
wb.save(excel_path)
wb.close()

with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a') as writer:
    monthly_achievement[['Month', 'Impressions', 'Clicks', 'Cost_USD', 'CTR', 'CPC',
                         'Achievement_%', 'Status', 'Cumulative_Clicks', 'Cumulative_Cost',
                         'Cumulative_CTR', 'Cumulative_Achievement_%',
                         'CTR_MoM_Change', 'Clicks_MoM_Change', 'Cost_MoM_Change']].to_excel(
        writer, sheet_name='Monthly Achievement', index=False)
    
    channel_trend_display_export = (channel_trend_display * 100).round(2)
    channel_trend_display_export.to_excel(writer, sheet_name='Channel Trend')
    
    channel_cost_share.to_excel(writer, sheet_name='Budget Efficiency', index=False)

print("✅ Excel updated with 3 new sheets")

print("\n" + "=" * 70)
print("ALL DONE! Files created:")
print("=" * 70)
print(f"📄 {excel_path}")
for i in range(1, 18):
    print(f"📊 {output_dir}\\{str(i).zfill(2)}_*.png")
