import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.patches import Patch
import matplotlib
import datetime
import numpy as np # Import numpy for nan checking

# Use a backend that supports interactive display if needed,
# or 'Agg' for saving files without showing a window.
try:
    matplotlib.use('TkAgg')
except ImportError:
    print("TkAgg backend not available, using default.")

# --- Configuration ---
# Load the roadmap data
file_path = "../LSEG AI Roadmap 2025.xlsx" # Make sure this path is correct
output_image_path = "roadmap_with_milestones_per_activity.png" # Optional: Path to save the plot

# Set up color mapping for activity statuses
status_colors = {
    "in progress": "green",
    "planned": "blue",
    "discussing use cases": "orange",
    "on hold": "grey",
    "completed": "#77dd77"
    # Add other statuses and their colors here
}

# Define the date range for the x-axis
x_min = pd.Timestamp('2025-03-01')
x_max = pd.Timestamp('2025-07-31')

# Milestone appearance (per activity)
milestone_color = 'red'
milestone_linestyle = '-' # Solid line for the marker on the bar
milestone_linewidth = 1.5
milestone_marker_height = 0.6 # How tall the marker line is relative to bar height
milestone_label_fontsize = 6 # Reduced font size (original 9 - 3 = 6)

# --- Data Loading and Cleaning ---
try:
    df = pd.read_excel(file_path, sheet_name=0, header=None)
except FileNotFoundError:
    print(f"Error: File not found at {file_path}")
    exit()

# Clean headers and structure
headers = df.iloc[0].fillna('') + df.iloc[1].fillna('')
df.columns = headers
roadmap = df[2:].copy()

# Rename columns for clarity and consistency (ensure NO spaces for itertuples)
rename_mapping = {
    'LSEG AI Enablement Roadmap 2025Area': 'Area',
    'Activity': 'Activity',  # Ensure this is correct
    'Start Date': 'StartDate', # Renamed to valid identifier
    'End Date': 'EndDate',     # Renamed to valid identifier
    'Owner': 'Owner',
    'Contributors': 'Contributors',
    'Comments': 'Comments',
    'Type': 'Type',
    'Status': 'Status',
    'Milestone 1 Date': 'Milestone1Date', # Renamed
    'Milestone 1 Name': 'Milestone1Name', # Renamed
    'Milestone 2 Date': 'Milestone2Date', # Renamed
    'Milestone 2 Name': 'Milestone2Name'  # Renamed
    # Add more milestone columns here if needed following the pattern
}
# Filter rename_mapping to only include columns actually present in the dataframe
actual_columns_to_rename = {k: v for k, v in rename_mapping.items() if k in roadmap.columns}
roadmap = roadmap.rename(columns=actual_columns_to_rename)

print("Columns after renaming:", list(roadmap.columns)) # Verify names like StartDate, EndDate etc.

# Drop rows where essential activity data is missing
essential_cols = ['Activity', 'StartDate', 'EndDate']
for col in essential_cols:
     if col not in roadmap.columns:
         print(f"Error: Essential column '{col}' not found after renaming. Exiting.")
         exit()
roadmap = roadmap.dropna(subset=essential_cols).copy() # Use copy after dropna


# Convert date columns to datetime objects using the NEW names
date_cols = ['StartDate', 'EndDate', 'Milestone1Date', 'Milestone2Date']
for col in date_cols:
    if col in roadmap.columns:
        roadmap[col] = pd.to_datetime(roadmap[col], errors='coerce') # 'coerce' turns errors into NaT
    else:
        # This is okay for milestone columns, just not essential ones
        print(f"Info: Date column '{col}' not found, skipping conversion.")


# Ensure Status column exists and handle potential missing values
if 'Status' not in roadmap.columns:
    print("Warning: 'Status' column not found. Using default color.")
    roadmap['Status'] = "planned" # Assign a default status
else:
    # Ensure lowercase and fill NaNs AFTER potential column creation
    roadmap['Status'] = roadmap['Status'].fillna("planned").astype(str).str.lower()


# Define milestone column pairs using the RENAMED column names
milestone_pairs = [
    ('Milestone1Date', 'Milestone1Name'),
    ('Milestone2Date', 'Milestone2Name')
    # Add more pairs here if you have Milestone 3, 4, etc. using renamed cols
]

# --- Plotting ---
fig, ax = plt.subplots(figsize=(16, max(8, len(roadmap) * 0.6))) # Adjusted figsize

# Sort data for plotting
sort_columns = []
if 'Type' in roadmap.columns:
    sort_columns.append('Type')
if 'StartDate' in roadmap.columns:
    sort_columns.append('StartDate')
else: # Should not happen due to earlier check, but defensive coding
     print("Error: StartDate column missing before sorting.")
     exit()

# Determine ascending order (sort by Type descending, then StartDate ascending)
ascending_order = [col != 'Type' for col in sort_columns]
roadmap_sorted = roadmap.sort_values(by=sort_columns, ascending=ascending_order).reset_index(drop=True)

print(f"Plotting {len(roadmap_sorted)} activities...")

# Background shading based on 'Type'
for i, row in enumerate(roadmap_sorted.itertuples(index=False)):
    # Safely get 'Type' attribute if it exists
    row_type = getattr(row, 'Type', None)

    # Check if Type is numeric before comparison
    try:
        numeric_type = pd.to_numeric(row_type, errors='coerce')
        if pd.isna(numeric_type):
             pass # Skip shading if Type is not numeric or missing
        elif numeric_type == 1:
            ax.axhspan(i - 0.5, i + 0.5, color="#dddddd", alpha=0.6, zorder=0)
        elif numeric_type == 2:
            ax.axhspan(i - 0.5, i + 0.5, color="#e0f2ff", alpha=0.7, zorder=0)
    except Exception as e:
        # Handle cases where 'Type' might cause other issues
        # print(f"Warning: Could not process Type '{row_type}' for shading row {i}: {e}")
        pass


# Plot Gantt bars and associated milestones
for i, row in enumerate(roadmap_sorted.itertuples(index=False)):

    # --- Plot Activity Bar ---
    # Ensure dates are valid Timestamps before proceeding
    if not isinstance(row.StartDate, pd.Timestamp) or not isinstance(row.EndDate, pd.Timestamp):
        print(f"Skipping row {i} ('{getattr(row, 'Activity', 'N/A')}') due to invalid date types: Start={type(row.StartDate)}, End={type(row.EndDate)}")
        continue
    if row.StartDate > row.EndDate:
        print(f"Skipping row {i} ('{getattr(row, 'Activity', 'N/A')}') due to Start Date after End Date.")
        continue

    # Clamp bar dates to the plot's visible range
    bar_start = max(row.StartDate, x_min)
    bar_end = min(row.EndDate, x_max)

    # Only plot if the bar has positive duration within the visible range
    if bar_end > bar_start:
        bar_duration_days = (bar_end - bar_start).days + 1 # Include end date
        bar_width_timedelta = pd.Timedelta(days=bar_duration_days)
        status_val = getattr(row, 'Status', 'planned') # Safely get Status
        color = status_colors.get(status_val, 'grey') # Default to grey

        # Add bar
        ax.barh(
            y=i,
            width=bar_width_timedelta,
            left=bar_start,
            color=color,
            edgecolor='black',
            height=0.5,
            zorder=2 # Ensure bars are above background shading
        )

        # --- Add Activity Labels (New Logic) ---
        label_text = str(getattr(row, 'Activity', ''))
        if label_text: # Only proceed if there is an activity name
            # --- Calculate estimated widths ---
            # Bar width in days (visible portion clamped to axes)
            # Use bar_duration_days calculated earlier (days between bar_start and bar_end)
            if bar_duration_days <= 0: # No visible bar part
                 pass # Don't draw label if bar isn't visible
            else:
                # Approx Pixels per character (adjust 7 or 8 based on font/visuals)
                pixels_per_char = 7
                estimated_text_pixel_width = len(label_text) * pixels_per_char

                # Axes width in pixels
                ax_pos = ax.get_position() # Gets bounding box [left, bottom, width, height] in figure coords
                ax_pixel_width = ax_pos.width * fig.get_figwidth() * fig.dpi

                # Total date range width in days displayed on the axis
                total_days_range = (x_max - x_min).days
                if total_days_range <= 0: total_days_range = 1 # Avoid division by zero

                # Visible bar width in pixels (approximate)
                bar_pixel_width_approx = (bar_duration_days / total_days_range) * ax_pixel_width

                # --- Decide placement ---
                # Add a small buffer (e.g., 10-15 pixels) to the text width for padding
                buffer_pixels = 15
                fits_inside = (estimated_text_pixel_width + buffer_pixels) < bar_pixel_width_approx

                if fits_inside:
                    # Place Inside (Centered Horizontally)
                    text_x_date = bar_start + (bar_end - bar_start) / 2
                    ax.text(text_x_date, i, label_text,
                            va='center', ha='center', # Center horizontally and vertically
                            fontsize=8, color='white', fontweight='bold',
                            clip_on=True, # Prevent text spilling out if calculation is slightly off
                            zorder=3)
                else:
                    # Place Outside (Left)
                    # Adjust the 'days' offset for spacing from the axis start
                    label_offset_days = pd.Timedelta(days=-1)
                    text_x_date = x_min - label_offset_days
                    ax.text(text_x_date, i, label_text,
                            va='center', ha='left', # Align right edge of text to the position
                            fontsize=8, color='black', fontweight='bold',
                            zorder=3)


    # --- Plot Milestones for THIS Activity ---
    for ms_date_col, ms_name_col in milestone_pairs:
        # Check if this row has data for this milestone pair
        if hasattr(row, ms_date_col) and hasattr(row, ms_name_col):
            ms_date = getattr(row, ms_date_col)
            ms_name = getattr(row, ms_name_col)

            # Validate the milestone data for this row
            # Check date is Timestamp, name is not null/nan, name is not 'n/a' (case-insensitive)
            if isinstance(ms_date, pd.Timestamp) and pd.notna(ms_name) and str(ms_name).strip().lower() != 'n/a':

                # Check if milestone date is within the plot's visible x-range
                if x_min <= ms_date <= x_max:
                    # 1. Plot vertical marker line on the activity row
                    marker_y_start = i - (0.5 * milestone_marker_height / 2) # Center marker vertically
                    marker_y_end = i + (0.5 * milestone_marker_height / 2)
                    ax.plot([ms_date, ms_date], [marker_y_start, marker_y_end],
                            color=milestone_color,
                            linestyle=milestone_linestyle,
                            linewidth=milestone_linewidth,
                            solid_capstyle='butt', # Make line ends flat
                            zorder=3) # Above bar, potentially below label

                    ax.plot(ms_date, i,  # x=date, y=row index (center)
                            marker='d',  # Diamond marker style
                            markersize=7,  # Adjust size if needed (e.g., 6, 8)
                            color=milestone_color,  # Use the defined milestone color
                            linestyle='None',  # No line connecting markers
                            zorder=4)  # Place diamond above vertical line, below text
                    # 3. Add milestone label above the marker
                    ax.text(ms_date, i + 0.65, # Position slightly above the bar's top edge
                            str(ms_name), # Ensure name is string
                            ha='center', # Center horizontally on the date
                            va='bottom', # Anchor bottom of text to the y-position
                            fontsize=milestone_label_fontsize, # Use smaller font size
                            color=milestone_color,
                            rotation=0, # Keep label horizontal
                            zorder=5) # Ensure text is on top of everything


# Draw 'Today' line
today = pd.Timestamp(datetime.date.today())
if x_min <= today <= x_max:
    ax.axvline(today, color='black', linewidth=1.5, linestyle=':', label='Today', zorder=4) # zorder places it above markers/bars but below text

    # Add "Today" text label near the bottom
    ax.text(today,                     # X-coordinate is the date 'today'
            -0.02,                     # Y-coordinate in axes fraction (0=bottom, 1=top). -0.04 is slightly below the axis.
            'Today',                   # The text to display
            ha='center',               # Center the text horizontally on the 'today' date
            va='top',                  # Align the top of the text to the specified y-coordinate
            fontsize=8,                # Smaller font size
            color='black',             # Text color
            rotation=0,                # Keep text horizontal
            transform=ax.get_xaxis_transform(), # Key part: Use data coords for X, axes coords for Y
            clip_on=False,             # Allow text to be drawn outside the main plot area if needed
            zorder=5)                  # Ensure it's drawn on top
# --- Formatting Axes and Legend ---
ax.set_yticks(range(len(roadmap_sorted)))
# Keep Y-axis ticks but remove labels (labels are now next to bars)
ax.set_yticklabels(["" for _ in range(len(roadmap_sorted))])

ax.xaxis.set_major_locator(mdates.MonthLocator())
ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))
ax.xaxis.set_minor_locator(mdates.WeekdayLocator(interval=1, byweekday=mdates.MO)) # Minor ticks for Mondays
ax.xaxis.grid(True, which='major', linestyle='-', linewidth='0.5', color='grey', zorder=1) # Major gridlines
ax.xaxis.grid(True, which='minor', linestyle=':', linewidth='0.5', color='lightgrey', zorder=1) # Minor gridlines

ax.set_xlim(x_min, x_max)
# ax.set_ylim automatically adjusts, but ensure it's reasonable
y_min_plot, y_max_plot = ax.get_ylim()
ax.set_ylim(y_max_plot, y_min_plot) # Invert y-axis AFTER getting limits

ax.set_title("LSEG AI Roadmap 2025", fontsize=16, fontweight='bold')
ax.set_xlabel("Timeline (2025)", fontsize=12)
ax.set_ylabel("") # Remove Y-axis label


# Create legend handles excluding the generic milestone line
legend_patches = [Patch(color=color, label=label.capitalize()) for label, color in status_colors.items() if label in roadmap_sorted['Status'].unique()]
if x_min <= today <= x_max:
     today_line = plt.Line2D([0], [0], color='black', linestyle=':', linewidth=1.5, label='Today')
     legend_patches.append(today_line)

# Add a legend entry for the *meaning* of the milestone marker/label
milestone_marker_legend = plt.Line2D([0], [0], color=milestone_color, linestyle=milestone_linestyle, lw=milestone_linewidth, label='Milestone', marker='|', markersize=8, markeredgecolor=milestone_color) # Use a marker in legend
legend_patches.append(milestone_marker_legend)


ax.legend(handles=legend_patches, loc='lower left', fontsize=10, framealpha=0.9)

# --- Adjust Layout and Display ---
# No need for extra bottom margin adjustment anymore
plt.tight_layout() # Adjust plot to prevent labels overlapping axes etc.

# Ensure grid lines are behind bars
ax.set_axisbelow(True)

try:
    plt.savefig(output_image_path, dpi=300) # Save the figure
    print(f"Roadmap saved to {output_image_path}")
except Exception as e:
    print(f"Error saving figure: {e}")

plt.show() # Display the plot