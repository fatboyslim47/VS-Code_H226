"""Open and load the most recent timesheet_export Excel file.

This script searches for timesheet_export (N).xlsx files in the current directory
and loads the one with the highest revision number into a pandas DataFrame.
Creates a plot of hours spent by month on specific tasks.
"""

# Standard library imports
import os
import re
import math
import calendar
import io
import shutil
import subprocess
from pathlib import Path

# Third-party imports
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.image as mpimg


# Set to True only when you want to regenerate the summary vertical bar chart.
GENERATE_VERTICAL_BAR_CHART = False

# Set to True to automatically render the Quarto HTML report after script output.
AUTO_RENDER_QMD_REPORT = True

# Merge historical task IDs into active task IDs before aggregation.
TASK_ID_MERGE_MAP = {
    91971495: 92739080,
}


def resolve_quarto_executable() -> str | None:
    """Return a usable Quarto executable path, or None if not found."""
    quarto_on_path = shutil.which('quarto')
    if quarto_on_path:
        return quarto_on_path

    default_windows_quarto = Path('C:/Users/mwoodmansee/AppData/Local/Programs/Quarto/bin/quarto.exe')
    if default_windows_quarto.exists():
        return str(default_windows_quarto)

    return None


def render_qmd_report(script_dir: Path) -> None:
    """Render the task sparkline QMD report into the output folder."""
    qmd_file = script_dir / 'task_sparkline_landscape_report.qmd'
    if not qmd_file.exists():
        print(f"Quarto report file not found, skipping render: {qmd_file}")
        return

    quarto_executable = resolve_quarto_executable()
    if not quarto_executable:
        print('Quarto executable not found. Skipping HTML render step.')
        return

    output_dir = script_dir / 'output'
    output_dir.mkdir(parents=True, exist_ok=True)

    render_cmd = [
        quarto_executable,
        'render',
        str(qmd_file),
        '--to',
        'html',
        '--output',
        'task_sparkline_landscape_report.html',
        '--output-dir',
        str(output_dir),
    ]

    print('\nRendering Quarto HTML report...')
    try:
        subprocess.run(render_cmd, check=True, cwd=str(script_dir))
        print(f"HTML report saved to: {output_dir / 'task_sparkline_landscape_report.html'}")
    except subprocess.CalledProcessError as error:
        print(f"Quarto render failed with exit code {error.returncode}.")


def cleanup_stale_task_images(script_dir: Path, latest_revision: int, keep_task_ids: list[int]) -> None:
    """Delete task sparkline PNGs for the current revision that are no longer in scope."""
    keep_set = set(int(task_id) for task_id in keep_task_ids)
    pattern = re.compile(r'^task_(\d+)_employee_sparklines_rev(\d+)\.png$', re.IGNORECASE)

    for image_file in script_dir.glob('task_*_employee_sparklines_rev*.png'):
        match = pattern.match(image_file.name)
        if not match:
            continue
        image_task_id = int(match.group(1))
        image_revision = int(match.group(2))
        if image_revision == latest_revision and image_task_id not in keep_set:
            image_file.unlink(missing_ok=True)
            print(f"Removed stale task sparkline image: {image_file}")


def format_month_label(value: str) -> str:
    """Convert numeric month values to month names when possible."""
    try:
        month_num = int(float(value))
        if 1 <= month_num <= 12:
            return calendar.month_name[month_num]
    except (TypeError, ValueError):
        pass
    return str(value)


def create_sparkline_compilation_image(
    sparkline_images: list[tuple[str, any]],
    script_dir: Path,
    latest_revision: int,
    output_prefix: str,
    compilation_title: str,
) -> None:
    """Create one combined image containing all sparkline plots for an entity type."""
    if not sparkline_images:
        print("No sparkline files available for compilation image.")
        return

    n_files = len(sparkline_images)
    n_cols = 2
    n_rows = math.ceil(n_files / n_cols)

    fig, axes = plt.subplots(n_rows, n_cols, figsize=(18, 6 * n_rows))
    if n_rows == 1:
        axes = [axes]

    # Flatten axes so we can index consistently.
    axes_flat = []
    for row_axes in axes:
        if isinstance(row_axes, (list, tuple)):
            axes_flat.extend(row_axes)
        else:
            axes_flat.extend(list(row_axes))

    for idx, (image_label, image_data) in enumerate(sparkline_images):
        ax = axes_flat[idx]
        ax.imshow(image_data)
        ax.axis('off')
        ax.set_title(image_label, fontsize=10)

    for idx in range(n_files, len(axes_flat)):
        axes_flat[idx].axis('off')

    fig.suptitle(compilation_title, fontsize=16, fontweight='bold')
    plt.tight_layout(rect=[0, 0, 1, 0.97])

    compilation_file = script_dir / f'{output_prefix}_employee_sparklines_compilation_rev{latest_revision}.png'
    plt.savefig(compilation_file, dpi=300, bbox_inches='tight')
    print(f"Compilation image saved to: {compilation_file}")
    plt.close(fig)


def create_entity_employee_sparkline_plots(
    df_employee_grouped: pd.DataFrame,
    employee_col: str,
    entity_id_col: str,
    entity_name_col: str,
    entity_ids: list[int],
    title_prefix_col: str,
    title_prefix_label: str,
    title_entity_label: str,
    output_prefix: str,
    bar_color: str,
    script_dir: Path,
    latest_revision: int,
) -> list[tuple[str, any]]:
    """Create one column-sparkline chart per entity id, with one row per employee."""
    if df_employee_grouped.empty:
        print("\nNo employee-level data available for sparkline plots.")
        return []

    output_images: list[tuple[str, any]] = []

    # Use a single normalized month key type so groupby/reindex align correctly.
    df_employee_grouped = df_employee_grouped.copy()
    df_employee_grouped['MonthKey'] = df_employee_grouped['Month'].astype(str)
    month_values = df_employee_grouped['MonthKey']
    month_numbers = pd.to_numeric(month_values, errors='coerce')

    if month_numbers.notna().all():
        month_map = pd.DataFrame({'month_str': month_values, 'month_num': month_numbers}).drop_duplicates()
        month_labels = (
            month_map.sort_values(['month_num', 'month_str'])['month_str']
            .drop_duplicates()
            .tolist()
        )
    else:
        month_dates = pd.to_datetime(month_values, errors='coerce')
        if month_dates.notna().any():
            month_map = pd.DataFrame({'month_str': month_values, 'month_dt': month_dates}).drop_duplicates()
            month_labels = (
                month_map.sort_values(['month_dt', 'month_str'])['month_str']
                .drop_duplicates()
                .tolist()
            )
        else:
            month_labels = sorted(month_values.unique())

    available_entity_ids = set(df_employee_grouped[entity_id_col].dropna().astype(int).tolist())
    missing_entity_ids = [entity_id for entity_id in entity_ids if entity_id not in available_entity_ids]
    if missing_entity_ids:
        print(f"\nWarning: Requested {entity_id_col} values not found in employee data: {missing_entity_ids}")

    for entity_id in entity_ids:
        entity_df = df_employee_grouped[df_employee_grouped[entity_id_col] == entity_id].copy()
        if entity_df.empty:
            continue

        entity_name = str(entity_df[entity_name_col].iloc[0])
        if title_prefix_col in entity_df.columns:
            title_prefix_values = sorted(entity_df[title_prefix_col].astype(str).unique())
            title_prefix = ' | '.join(title_prefix_values)
        else:
            title_prefix = f'Unknown {title_prefix_label}'

        employees = sorted(entity_df[employee_col].astype(str).unique())
        n_rows = len(employees)
        fig_height = max(2.5, (0.55 * n_rows) + 1.8)

        fig, axes = plt.subplots(n_rows, 1, figsize=(12, fig_height), sharex=True)
        if n_rows == 1:
            axes = [axes]

        # For visual comparison across employees on the same entity, use one shared Y-scale.
        if n_rows > 1:
            shared_entity_max_hours = max(float(entity_df['Hours (h)'].max()), 1.0)
        else:
            shared_entity_max_hours = None

        for idx, employee in enumerate(employees):
            ax = axes[idx]
            emp_df = entity_df[entity_df[employee_col].astype(str) == employee]
            emp_series = (
                emp_df.groupby('MonthKey')['Hours (h)']
                .sum()
                .reindex(month_labels, fill_value=0)
            )
            employee_total_hours = float(emp_series.sum())
            employee_max_month_hours = max(float(emp_series.max()), 1.0)
            y_axis_max = shared_entity_max_hours if shared_entity_max_hours is not None else employee_max_month_hours

            ax.bar(range(len(month_labels)), emp_series.values, color=bar_color, width=0.8)
            ax.set_ylim(0, y_axis_max * 1.15)
            ax.set_yticks([])
            ax.grid(axis='y', alpha=0.2, linewidth=0.5)
            # Place employee labels outside the plot area to avoid overlap with first bar.
            ax.text(-0.08, 0.82, employee, transform=ax.transAxes, fontsize=8, va='top', ha='right', clip_on=False)

            # Label each monthly bar with hours above the column.
            for x_pos, hour_value in enumerate(emp_series.values):
                if hour_value > 0:
                    ax.text(
                        x_pos,
                        float(hour_value) + (y_axis_max * 0.03),
                        f"{float(hour_value):.1f}",
                        ha='center',
                        va='bottom',
                        fontsize=7,
                        color='#264653',
                    )

            ax.text(
                -0.08,
                0.56,
                f"Total: {employee_total_hours:.1f}h",
                transform=ax.transAxes,
                fontsize=8,
                va='top',
                ha='right',
                clip_on=False,
            )

            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_visible(False)
            if idx < n_rows - 1:
                ax.spines['bottom'].set_visible(False)
                ax.tick_params(axis='x', which='both', bottom=False, labelbottom=False)

        month_display_labels = [format_month_label(month) for month in month_labels]
        axes[-1].set_xticks(range(len(month_labels)))
        axes[-1].set_xticklabels(month_display_labels, rotation=45, ha='right', fontsize=8)

        if str(title_entity_label).lower() == 'task':
            # Ensure task plots are titled as: Project, Task Name, then Task ID.
            if str(title_prefix) and str(title_prefix) != str(entity_name):
                sparkline_title = f'{title_prefix} | {entity_name} | Task ID {entity_id}'
            else:
                sparkline_title = f'{entity_name} | Task ID {entity_id}'
        else:
            if str(title_prefix) == str(entity_name):
                sparkline_title = f'{title_prefix}'
            else:
                sparkline_title = f'{title_prefix} | {entity_name}'

        fig.suptitle(
            sparkline_title,
            fontsize=12,
            fontweight='bold',
        )
        plt.tight_layout(rect=[0.16, 0, 1, 0.95])

        image_label = f'{output_prefix}_{entity_id}_employee_sparklines_rev{latest_revision}'
        image_file = script_dir / f'{image_label}.png'
        fig.savefig(image_file, format='png', dpi=300, bbox_inches='tight')
        print(f"Saved sparkline image: {image_file}")

        buffer = io.BytesIO()
        fig.savefig(buffer, format='png', dpi=300, bbox_inches='tight')
        buffer.seek(0)
        output_images.append((image_label, mpimg.imread(buffer)))
        buffer.close()
        plt.close(fig)

    return output_images

# ---------------------------------------------------------------------------
# Find timesheet_export files
# ---------------------------------------------------------------------------

# Get the directory where this script is located
script_dir = Path(__file__).parent

# Pattern to match: timesheet_export (N).xlsx where N is the revision number
pattern = re.compile(r'^timesheet_export \((\d+)\)\.xlsx$', re.IGNORECASE)

# Find all matching files and extract revision numbers
matching_files = []
for file in script_dir.glob('timesheet_export*.xlsx'):
    match = pattern.match(file.name)
    if match:
        revision_number = int(match.group(1))
        matching_files.append((revision_number, file))

# Check if any files were found
if not matching_files:
    raise FileNotFoundError("No timesheet_export (N).xlsx files found in the directory.")

# Sort by revision number and get the most recent (highest number)
matching_files.sort(key=lambda x: x[0], reverse=True)
latest_revision, latest_file = matching_files[0]

print(f"Found {len(matching_files)} timesheet_export file(s)")
print(f"Opening latest revision: {latest_file.name} (revision {latest_revision})")

# ---------------------------------------------------------------------------
# Load the Excel file
# ---------------------------------------------------------------------------

# Read the Excel file into a pandas DataFrame
df = pd.read_excel(latest_file, header=0)

print(f"\nLoaded {len(df)} rows and {len(df.columns)} columns")
print(f"\nColumns: {list(df.columns)}")
print(f"\nFirst few rows:")
print(df.head())

# Display basic info about the DataFrame
print(f"\nDataFrame info:")
df.info()

# ---------------------------------------------------------------------------
# Curate data for plotting hours by month and task
# ---------------------------------------------------------------------------

# Select relevant columns for analysis
required_columns = ['Month', 'Task Id', 'Task', 'Project Id', 'Project', 'Hours (h)']
employee_column_candidates = ['Employee', 'Employee Name', 'Assignee', 'User', 'Person', 'Resource', 'Member', 'For']
employee_col = next((col for col in employee_column_candidates if col in df.columns), None)

if employee_col is None:
    print("\nWarning: No employee column found. Sparkline-by-employee plots will be skipped.")
else:
    print(f"\nUsing employee column for sparklines: {employee_col}")

# Check if all required columns exist
missing_columns = [col for col in required_columns if col not in df.columns]
if missing_columns:
    print(f"\nWarning: Missing columns: {missing_columns}")
    print("Available columns:", list(df.columns))
else:
    # Filter to only the columns we need
    selected_columns = required_columns + ([employee_col] if employee_col else [])
    df_curated = df[selected_columns].copy()
    
    # Remove rows with missing values in key columns
    dropna_subset = ['Month', 'Task Id', 'Task', 'Project Id', 'Project', 'Hours (h)'] + ([employee_col] if employee_col else [])
    df_curated = df_curated.dropna(subset=dropna_subset)

    # Merge legacy Task IDs into active Task IDs so reporting rolls into one task.
    if TASK_ID_MERGE_MAP:
        task_name_by_id = (
            df_curated[['Task Id', 'Task']]
            .drop_duplicates(subset=['Task Id'])
            .set_index('Task Id')['Task']
            .to_dict()
        )
        task_project_by_id = (
            df_curated[['Task Id', 'Project Id', 'Project']]
            .drop_duplicates(subset=['Task Id'])
            .set_index('Task Id')[['Project Id', 'Project']]
            .to_dict('index')
        )
        df_curated['Task Id'] = df_curated['Task Id'].astype(int)
        for old_task_id, new_task_id in TASK_ID_MERGE_MAP.items():
            old_rows = df_curated['Task Id'] == old_task_id
            if old_rows.any():
                df_curated.loc[old_rows, 'Task Id'] = new_task_id
                replacement_name = task_name_by_id.get(new_task_id)
                if replacement_name:
                    df_curated.loc[df_curated['Task Id'] == new_task_id, 'Task'] = replacement_name

                replacement_project = task_project_by_id.get(new_task_id)
                if replacement_project:
                    df_curated.loc[df_curated['Task Id'] == new_task_id, 'Project Id'] = replacement_project['Project Id']
                    df_curated.loc[df_curated['Task Id'] == new_task_id, 'Project'] = replacement_project['Project']
                print(f"Merged Task ID {old_task_id} into Task ID {new_task_id}")
    
    # Group by Month, Task, and Project, summing the hours
    df_grouped = df_curated.groupby(['Month', 'Task Id', 'Task', 'Project Id', 'Project'])['Hours (h)'].sum().reset_index()
    
    print(f"\nCurated data summary:")
    print(f"Total tasks: {df_grouped['Task'].nunique()}")
    print(f"Total months: {df_grouped['Month'].nunique()}")
    print(f"\nGrouped data preview:")
    print(df_grouped.head(10))
    
    # -----------------------------------------------------------------------
    # Task Id aggregation
    # -----------------------------------------------------------------------

    # Filter to specific Task Ids
    specific_task_ids = [66391038, 94452732, 82982100, 92780832, 94453589, 94279175, 92739080, 93076386]
    df_filtered = df_grouped[df_grouped['Task Id'].isin(specific_task_ids)].copy()
    
    print(f"\nFiltered data summary (specific Task Ids):")
    print(f"Total tasks included: {df_filtered['Task'].nunique()}")
    print(f"Total months: {df_filtered['Month'].nunique()}")
    print(f"\nFiltered data preview:")
    print(df_filtered.head(10))
    missing_task_ids = sorted(set(specific_task_ids) - set(df_filtered['Task Id'].dropna().astype(int).tolist()))
    if missing_task_ids:
        print(f"Warning: Requested Task Ids not found in grouped data: {missing_task_ids}")

    if employee_col:
        df_task_employee_grouped = (
            df_curated.groupby(['Month', 'Task Id', 'Task', 'Project', employee_col])['Hours (h)']
            .sum()
            .reset_index()
        )
        df_task_employee_filtered = df_task_employee_grouped[
            df_task_employee_grouped['Task Id'].isin(specific_task_ids)
        ].copy()
        task_sparkline_images = create_entity_employee_sparkline_plots(
            df_employee_grouped=df_task_employee_filtered,
            employee_col=employee_col,
            entity_id_col='Task Id',
            entity_name_col='Task',
            entity_ids=specific_task_ids,
            title_prefix_col='Project',
            title_prefix_label='Project',
            title_entity_label='Task',
            output_prefix='task',
            bar_color='#2a9d8f',
            script_dir=script_dir,
            latest_revision=latest_revision,
        )
        create_sparkline_compilation_image(
            sparkline_images=task_sparkline_images,
            script_dir=script_dir,
            latest_revision=latest_revision,
            output_prefix='all_task',
            compilation_title='Employee Sparkline Compilation - All Tasks',
        )
        cleanup_stale_task_images(
            script_dir=script_dir,
            latest_revision=latest_revision,
            keep_task_ids=specific_task_ids,
        )

    # -----------------------------------------------------------------------
    # Project Id aggregation
    # -----------------------------------------------------------------------

    specific_project_ids = [92723243]
    df_project_source = df_curated[df_curated['Project Id'].isin(specific_project_ids)].copy()
    df_project_grouped = (
        df_project_source.groupby(['Month', 'Project Id', 'Project'])['Hours (h)']
        .sum()
        .reset_index()
    )
    df_project_filtered = df_project_grouped.copy()

    print(f"\nFiltered data summary (specific Project Ids):")
    print(f"Raw rows matched on Project Id column: {len(df_project_source)}")
    print(f"Total projects included: {df_project_filtered['Project'].nunique()}")
    print(f"Total months: {df_project_filtered['Month'].nunique()}")
    print(f"\nFiltered Project data preview:")
    print(df_project_filtered.head(10))
    missing_project_ids = sorted(set(specific_project_ids) - set(df_project_source['Project Id'].dropna().astype(int).tolist()))
    if missing_project_ids:
        print(f"Warning: Requested Project Ids not found in grouped data: {missing_project_ids}")

    if employee_col:
        df_project_employee_grouped = (
            df_project_source.groupby(['Month', 'Project Id', 'Project', employee_col])['Hours (h)']
            .sum()
            .reset_index()
        )
        project_sparkline_images = create_entity_employee_sparkline_plots(
            df_employee_grouped=df_project_employee_grouped,
            employee_col=employee_col,
            entity_id_col='Project Id',
            entity_name_col='Project',
            entity_ids=specific_project_ids,
            title_prefix_col='Project',
            title_prefix_label='Project',
            title_entity_label='Project',
            output_prefix='project',
            bar_color='#1d4ed8',
            script_dir=script_dir,
            latest_revision=latest_revision,
        )
        create_sparkline_compilation_image(
            sparkline_images=project_sparkline_images,
            script_dir=script_dir,
            latest_revision=latest_revision,
            output_prefix='all_project',
            compilation_title='Employee Sparkline Compilation - All Projects',
        )
    
    # ---------------------------------------------------------------------------
    # Optional vertical bar chart output
    # ---------------------------------------------------------------------------

    if GENERATE_VERTICAL_BAR_CHART:
        # Pivot the data to have months as rows and tasks as columns
        df_pivot = df_filtered.pivot_table(index='Month', columns=['Task Id', 'Task', 'Project Id', 'Project'], values='Hours (h)', aggfunc='sum')

        # Create the plot
        fig, ax = plt.subplots(figsize=(12, 8))
        df_pivot.plot(kind='bar', stacked=True, ax=ax)

        # Replace numeric month ticks with month names for readability.
        x_labels = [format_month_label(str(idx)) for idx in df_pivot.index]
        ax.set_xticklabels(x_labels, rotation=45, ha='right')

        ax.set_xlabel('Month', fontsize=12)
        ax.set_ylabel('Hours (h)', fontsize=12)
        ax.set_title('Hours Spent by Month on Specific Tasks', fontsize=14, fontweight='bold')
        ax.legend(title='Task', bbox_to_anchor=(1.05, 1), loc='upper left')
        ax.grid(True, alpha=0.3)

        plt.tight_layout()

        # Save the plot
        output_file = script_dir / f'hours_by_month_plot_rev{latest_revision}.png'
        plt.savefig(output_file, dpi=300, bbox_inches='tight')
        print(f"\nPlot saved to: {output_file}")
        plt.close(fig)
    else:
        print("\nVertical bar chart generation is disabled.")
    
    # Optionally save the curated data to CSV
    output_csv = script_dir / f'hours_by_month_data_rev{latest_revision}.csv'
    df_filtered.to_csv(output_csv, index=False)
    print(f"Curated data saved to: {output_csv}")

    if AUTO_RENDER_QMD_REPORT:
        render_qmd_report(script_dir)
