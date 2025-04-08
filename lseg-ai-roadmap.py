import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.patches import Patch
import matplotlib
import datetime
matplotlib.use('TkAgg')




# Load the roadmap data
file_path = "../LSEG AI Roadmap 2025.xlsx"
df = pd.read_excel(file_path, sheet_name=0, header=None)

# Clean headers and structure
headers = df.iloc[0].fillna('') + df.iloc[1].fillna('')
df.columns = headers
roadmap = df[2:].copy()
roadmap = roadmap.rename(columns={
    'LSEG AI Enablement Roadmap 2025Area': 'Area',
    'Acitivity': 'Activity',
    'Start Date': 'Start Date',
    'End Date': 'End Date',
    'Owner': 'Owner',
    'Contributors': 'Contributors',
    'Comments': 'Comments',
    'Type': 'Type'
})

# Drop empty activity rows
roadmap = roadmap[roadmap['Activity'].notna()]
roadmap['Start Date']=pd.to_datetime(roadmap['Start Date'], errors='coerce')
roadmap['End Date'] = pd.to_datetime(roadmap['End Date'], errors='coerce')



roadmap['Status'] = roadmap['Status']

# Set up color mapping
status_colors = {
    "in progress": "green",
    "planned": "blue",
    "discussing use cases": "orange",
    "on hold": "grey"
}
x_min = pd.Timestamp('2025-03-01')
x_max = pd.Timestamp('2025-07-31')
# Plot Gantt chart
fig, ax = plt.subplots(figsize=(14, len(roadmap) * 0.6))
roadmap_sorted = roadmap.sort_values(by='Type', ascending=False).reset_index(drop=True)

for i, row in enumerate(roadmap_sorted.itertuples()):
    row_type = getattr(row, 'Type', None)

    if row_type == 1:
        ax.axhspan(i - 0.5, i + 0.5, color="#aaaaaa", zorder=0)  # light gray
    elif row_type == 2:
        ax.axhspan(i - 0.5, i + 0.5, color="#e0f2ff", zorder=0)  # light blue
for i, row in enumerate(roadmap_sorted.itertuples()):
    ax.barh(
        y=i,
        width=(row._4 - row._3).days,
        left=row._3,
        color=status_colors.get(row.Status, 'black'),
        edgecolor='black',
        height=0.5
    )
    bar_start = max(row._3, x_min)  # Clamp to March 1
    bar_end = min(row._4, x_max)  # Clamp to July 31
    bar_duration = (bar_end - bar_start).days
    bar_duration = (bar_end - bar_start).days
    label_text = str(row.Activity)
    color = status_colors.get(row.Status, 'black')

    # Add bar
    ax.barh(
        y=i,
        width=bar_duration,
        left=bar_start,
        color=color,
        edgecolor='black',
        height=0.5
    )

    # Estimate label width (e.g., 6.5px per character in horizontal space)
    label_char_width = 6.5
    max_label_width = bar_duration * 1.5  # adjust multiplier to tweak behavior

    # Choose label placement
    if len(label_text) * label_char_width < bar_duration*5:
        ax.text(bar_start, i, "  "+label_text, va='center', ha='left', fontsize=8, color='white', fontweight='bold')
    else:
        ax.text(bar_start - pd.Timedelta(days=1), i, label_text, va='center', ha='right', fontsize=8, color='black',
                fontweight='bold')

# Draw a thin vertical black line at today's date
today = pd.Timestamp(datetime.date.today())
ax.axvline(today, color='black', linewidth=1, linestyle='--', label='Today')

# Format axes
ax.set_yticks(range(len(roadmap)))
ax.set_yticklabels(["" for _ in range(len(roadmap))])
ax.xaxis.set_major_locator(mdates.MonthLocator())
ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))

ax.set_xlim(x_min, x_max)
ax.set_title("LSEG AI Roadmap 2025", fontsize=14)
ax.set_xlabel("Timeline (2025)")

# Legend
legend_patches = [Patch(color=color, label=label) for label, color in status_colors.items()]
ax.legend(handles=legend_patches, loc='lower left')

plt.tight_layout()
plt.show()
