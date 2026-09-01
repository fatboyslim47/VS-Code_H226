"""Create the Defense project sparkline report from the latest timesheet export.

This script searches for timesheet_export (N).xlsx files in the current directory
and loads the one with the highest revision number into a pandas DataFrame.
Creates employee-by-month sparkline plots for specific projects.
"""

# Standard library imports
import re
import math
import calendar
import io
import os
import shutil
import subprocess
import textwrap
from datetime import datetime
from pathlib import Path
from typing import Any
from matplotlib.backends.backend_pdf import PdfPages

# Third-party imports
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.image as mpimg


# Set to True only when you want to regenerate the summary vertical bar chart.
GENERATE_VERTICAL_BAR_CHART = False

# Set to True to automatically render the Quarto HTML report after script output.
AUTO_RENDER_QMD_REPORT = True

# The compilation image is not used by the HTML/PDF and can consume substantial memory.
GENERATE_COMPILATION_IMAGE = False

# Projects below this total are excluded from the project table and dedicated pages.
MIN_REPORT_PROJECT_HOURS = 10.0

# PDF page size: 11x17" paper in landscape orientation (width x height, inches).
PDF_PAGE_SIZE_INCHES = (17, 11)
# PDF page size: 11x17" paper in portrait orientation, used for specific pages below.
PDF_PAGE_SIZE_PORTRAIT_INCHES = (11, 17)

# Project/Task pages rendered in portrait rather than landscape orientation.
PDF_PORTRAIT_PROJECT_IDS: set[int] = set()
PDF_PORTRAIT_TASK_IDS: set[int] = set()

# Merge historical task IDs into active task IDs before aggregation.
TASK_ID_MERGE_MAP: dict[int, int] = {}

REQUIRED_COLUMNS = ['Month', 'Task Id', 'Task', 'Project Id', 'Project', 'Hours (h)']
EMPLOYEE_COLUMN_CANDIDATES = ['Employee', 'Employee Name', 'Assignee', 'User', 'Person', 'Resource', 'Member', 'For']
EMPLOYEE_NUMBER_COLUMN = 'Employee Number (Profile)'
DEFENSE_EMPLOYEE_NUMBER_BY_LP_ID = {
    264082: 5514,  # Andrew Clayton
    132491: 5253,  # Brock Francis
    111714: 5164,  # Donald Gill Jr.
    132472: 5433,  # Benjamin Groleau
    128148: 5198,  # Michael Johnson
    132473: 5392,  # Christopher Matthews
    211713: 5434,  # Colin McRae
    207441: 5420,  # Gregory Norris
    207433: 5124,  # Jerrald Ogle
    255931: 5500,  # Rushton Westcott
    263422: 5511,  # David Wolniewicz
}
# No Defense task-level breakdown project has been selected yet.
TASK_BREAKDOWN_PROJECT_ID: int | None = None
TASK_BAR_COLOR = 'green'

PROJECT_CATEGORY_BY_ID: dict[int, str] = {}

PROJECT_REFERENCE_BY_ID: dict[int, str] = {}

# Defense projects are discovered from employee time in each input workbook.
ADMIN_PROJECT_IDS = {
    58589966,
    57916280,
    57916278,
    58068435,
    61538760,
    57916281,
}
SUPPORT_PROJECT_IDS = {
    66389248,
    75709273,
    57163983,
    51842049,
    75709167,
    52104089,
}
CATEGORY_ORDER = ['ADMIN', 'NRE', 'SUPPORT']
CATEGORY_COLORS = {
    'ADMIN': '#577590',
    'NRE': '#1d4ed8',
    'SUPPORT': '#2a9d8f',
}
PROJECT_HOUR_TARGETS: dict[int, dict[str, float | None]] = {}


def resolve_quarto_executable() -> str | None:
    """Return a usable Quarto executable path, or None if not found."""
    quarto_on_path = shutil.which('quarto')
    if quarto_on_path:
        return quarto_on_path

    default_windows_quarto = Path('C:/Users/mwoodmansee/AppData/Local/Programs/Quarto/bin/quarto.exe')
    if default_windows_quarto.exists():
        return str(default_windows_quarto)

    return None


def render_qmd_report(script_dir: Path, output_prefix: str, executed_at: str) -> bool:
    """Render the project sparkline QMD report into the output folder."""
    qmd_file = script_dir / 'Defense_LP_Project-plotting.qmd'
    if not qmd_file.exists():
        print(f"Quarto report file not found, skipping render: {qmd_file}")
        return False

    quarto_executable = resolve_quarto_executable()
    if not quarto_executable:
        print('Quarto executable not found. Skipping HTML render step.')
        return False

    output_dir = script_dir / 'output'
    output_dir.mkdir(parents=True, exist_ok=True)
    html_output = output_dir / f'{output_prefix}.html'
    rendered_html = script_dir / f'{output_prefix}.html'

    # Remove old artifacts first so a failed render cannot leave stale HTML in place.
    html_output.unlink(missing_ok=True)
    rendered_html.unlink(missing_ok=True)

    render_cmd = [
        quarto_executable,
        'render',
        str(qmd_file),
        '--to',
        'html',
        '--output',
        f'{output_prefix}.html',
    ]

    print('\nRendering Quarto HTML report...')
    try:
        render_env = dict(
            **os.environ,
            LP_REPORT_ROOT=str(script_dir),
            LP_REPORT_PREFIX=output_prefix,
            LP_REPORT_EXECUTED_AT=executed_at,
        )
        subprocess.run(
            render_cmd,
            check=True,
            cwd=str(script_dir),
            env=render_env,
            capture_output=True,
            text=True,
        )
        if rendered_html.exists():
            html_output.unlink(missing_ok=True)
            shutil.move(str(rendered_html), str(html_output))
        if html_output.exists():
            print(f"HTML report saved to: {html_output}")
            return True
        print(f"Quarto completed, but HTML output was not found: {html_output}")
        return False
    except subprocess.CalledProcessError as error:
        print(f"Quarto render failed with exit code {error.returncode}.")
        if error.stdout:
            print(error.stdout)
        if error.stderr:
            print(error.stderr)
        return False


def delete_all_png_files(script_dir: Path) -> None:
    """Delete all PNG files in the script directory after successful HTML render."""
    png_files = sorted(script_dir.glob('*.png'))
    if not png_files:
        print('No PNG files found to delete.')
        return

    for png_file in png_files:
        png_file.unlink(missing_ok=True)
        print(f'Deleted PNG file: {png_file}')


def generate_project_pdf_report(
    script_dir: Path,
    latest_revision: int,
    output_prefix: str,
    executed_at: str,
    ordered_project_ids: list[int],
    task_breakdown_project_id: int | None = None,
    ordered_task_ids: list[int] | None = None,
    task_name_by_id: dict[int, str] | None = None,
    project_total_hours_by_id: dict[int, float] | None = None,
    task_total_hours_by_id: dict[int, float] | None = None,
    project_month_hours_by_id: dict[int, dict[str, float]] | None = None,
    project_month_labels: list[str] | None = None,
) -> bool:
    """Generate the mixed-orientation PDF report with totals in page titles."""
    output_dir = script_dir / 'output'
    output_dir.mkdir(parents=True, exist_ok=True)
    pdf_file = output_dir / f'{output_prefix}.pdf'

    image_pattern = re.compile(rf'^{re.escape(output_prefix)}_project_(\d+)_employee_sparklines_rev(\d+)\.png$', re.IGNORECASE)
    latest_images_by_project: dict[int, Path] = {}
    for image_file in script_dir.glob(f'{output_prefix}_project_*_employee_sparklines_rev*.png'):
        match = image_pattern.match(image_file.name)
        if not match:
            continue
        project_id = int(match.group(1))
        revision = int(match.group(2))
        if revision == latest_revision:
            latest_images_by_project[project_id] = image_file

    ordered_images: list[tuple[int, Path]] = [
        (project_id, latest_images_by_project[project_id])
        for project_id in ordered_project_ids
        if project_id in latest_images_by_project
    ]

    task_image_pattern = re.compile(rf'^{re.escape(output_prefix)}_task_(\d+)_employee_sparklines_rev(\d+)\.png$', re.IGNORECASE)
    latest_images_by_task: dict[int, Path] = {}
    if ordered_task_ids:
        for image_file in script_dir.glob(f'{output_prefix}_task_*_employee_sparklines_rev*.png'):
            match = task_image_pattern.match(image_file.name)
            if not match:
                continue
            task_id = int(match.group(1))
            revision = int(match.group(2))
            if revision == latest_revision:
                latest_images_by_task[task_id] = image_file

    category_pages: list[tuple[str, Path]] = []
    hours_chart = script_dir / f'{output_prefix}_category_hours_by_month_rev{latest_revision}.png'
    if hours_chart.exists():
        category_pages.append(('Monthly Hours by Category', hours_chart))
    percent_chart = script_dir / f'{output_prefix}_category_percent_by_month_rev{latest_revision}.png'
    if percent_chart.exists():
        category_pages.append(('Monthly Category Share (100% Stacked)', percent_chart))

    if not category_pages and not ordered_images:
        print('No category or project PNG files found for PDF generation.')
        return False

    def add_image_page(image_file: Path, title: str, page_size: tuple[int, int]) -> None:
        nonlocal page_num
        image_data = mpimg.imread(image_file)
        is_portrait = page_size == PDF_PAGE_SIZE_PORTRAIT_INCHES
        title_width = 72 if is_portrait else 110
        wrapped_title = '\n'.join(
            textwrap.wrap(
                title,
                width=title_width,
                break_long_words=False,
                break_on_hyphens=False,
            )
        )
        title_line_count = wrapped_title.count('\n') + 1
        image_top = 0.88 if title_line_count > 1 else 0.92
        fig = plt.figure(figsize=page_size)
        ax = fig.add_axes([0.04, 0.12, 0.92, image_top - 0.12])
        ax.imshow(image_data)
        ax.axis('off')
        fig.suptitle(
            wrapped_title,
            fontsize=12,
            fontweight='bold',
            y=0.97,
            va='top',
            linespacing=1.15,
        )
        fig.text(0.5, 0.03, f'Page {page_num} | Executed: {executed_at}', ha='center', va='center', fontsize=9)
        page_orientation = 'portrait' if is_portrait else 'landscape'
        pdf.savefig(fig, orientation=page_orientation)
        plt.close(fig)
        page_num += 1

    try:
        with PdfPages(pdf_file) as pdf:
            page_num = 1

            # Put the ordered project reference table on the first PDF page.
            month_labels = project_month_labels or []
            table_columns = ['Category', 'Project', 'Project ID'] + month_labels + ['Total Hours']
            table_rows = []
            for project_id in ordered_project_ids:
                monthly_hours = (project_month_hours_by_id or {}).get(project_id, {})
                table_rows.append([
                    PROJECT_CATEGORY_BY_ID.get(project_id, 'Unknown'),
                    PROJECT_REFERENCE_BY_ID.get(project_id, 'Unknown Project'),
                    str(project_id),
                    *[f'{monthly_hours.get(month_label, 0.0):.1f}' for month_label in month_labels],
                    f'{(project_total_hours_by_id or {}).get(project_id, 0.0):.1f}',
                ])

            table_fig = plt.figure(figsize=PDF_PAGE_SIZE_INCHES)
            table_ax = table_fig.add_axes([0.02, 0.08, 0.96, 0.84])
            table_ax.axis('off')
            table_ax.set_title(
                'Defense Engineering | Project Time Allocation',
                fontsize=16,
                fontweight='bold',
                pad=18,
            )
            table = table_ax.table(
                cellText=table_rows,
                colLabels=table_columns,
                colWidths=[0.11, 0.38, 0.09, *([0.04] * len(month_labels)), 0.085],
                cellLoc='center',
                colLoc='center',
                loc='center',
            )
            table.auto_set_font_size(False)
            table.set_fontsize(6.5)
            table.scale(1, 1.55)
            for cell in table.get_celld().values():
                cell.set_edgecolor('#B8C2CC')
            for column_index in range(len(table_columns)):
                cell = table[0, column_index]
                cell.set_facecolor('#264653')
                cell.set_text_props(color='white', weight='bold')
            total_hours_column_index = len(table_columns) - 1
            for row_index in range(1, len(table_rows) + 1):
                table[row_index, 1].set_text_props(weight='bold')
                table[row_index, total_hours_column_index].set_text_props(weight='bold')
            table_fig.text(0.5, 0.03, f'Page {page_num} | Executed: {executed_at}', ha='center', va='center', fontsize=9)
            pdf.savefig(table_fig, orientation='landscape')
            plt.close(table_fig)
            page_num += 1

            for page_title, image_file in category_pages:
                add_image_page(image_file, page_title, PDF_PAGE_SIZE_INCHES)

            for project_id, image_file in ordered_images:
                project_name = PROJECT_REFERENCE_BY_ID.get(project_id, 'Unknown Project')
                project_total = (project_total_hours_by_id or {}).get(project_id, 0.0)
                add_image_page(
                    image_file,
                    f'{project_name} | Project ID {project_id} | Total: {project_total:.1f}h',
                    PDF_PAGE_SIZE_PORTRAIT_INCHES,
                )

                # Insert the Task Id breakdown pages directly below their parent project page.
                if project_id == task_breakdown_project_id and ordered_task_ids:
                    for task_id in ordered_task_ids:
                        task_image_file = latest_images_by_task.get(task_id)
                        if not task_image_file:
                            continue
                        task_name = (task_name_by_id or {}).get(task_id, 'Unknown Task')
                        task_total = (task_total_hours_by_id or {}).get(task_id, 0.0)
                        add_image_page(
                            task_image_file,
                            f'{task_name} | Task ID {task_id} | Total: {task_total:.1f}h',
                            PDF_PAGE_SIZE_PORTRAIT_INCHES,
                        )

        print(f'PDF report saved to: {pdf_file}')
        return True
    except Exception as error:
        print(f'Failed to generate PDF report: {error}')
        return False


def cleanup_stale_entity_images(
    script_dir: Path,
    latest_revision: int,
    keep_entity_ids: list[int],
    output_prefix: str,
    file_prefix: str,
) -> None:
    """Delete entity sparkline PNGs for the current revision that are no longer in scope."""
    keep_set = set(int(entity_id) for entity_id in keep_entity_ids)
    pattern = re.compile(
        rf'^{re.escape(file_prefix)}_{re.escape(output_prefix)}_(\d+)_employee_sparklines_rev(\d+)\.png$',
        re.IGNORECASE,
    )

    for image_file in script_dir.glob(f'{file_prefix}_{output_prefix}_*_employee_sparklines_rev*.png'):
        match = pattern.match(image_file.name)
        if not match:
            continue
        image_entity_id = int(match.group(1))
        image_revision = int(match.group(2))
        if image_revision == latest_revision and image_entity_id not in keep_set:
            image_file.unlink(missing_ok=True)
            print(f"Removed stale {output_prefix} sparkline image: {image_file}")


def format_month_label(value: str) -> str:
    """Convert numeric month values to month names when possible."""
    try:
        month_num = int(float(value))
        if 1 <= month_num <= 12:
            return calendar.month_name[month_num]
    except (TypeError, ValueError):
        pass
    return str(value)


def find_latest_timesheet_export(script_dir: Path) -> tuple[int, Path, int]:
    """Find the latest timesheet export revision file in the script directory."""
    pattern = re.compile(r'^timesheet_export \((\d+)\)\.xlsx$', re.IGNORECASE)
    matching_files: list[tuple[int, Path]] = []

    for file in script_dir.glob('timesheet_export*.xlsx'):
        match = pattern.match(file.name)
        if match:
            revision_number = int(match.group(1))
            matching_files.append((revision_number, file))

    if not matching_files:
        raise FileNotFoundError('No timesheet_export (N).xlsx files found in the directory.')

    matching_files.sort(key=lambda x: x[0], reverse=True)
    latest_revision, latest_file = matching_files[0]
    return latest_revision, latest_file, len(matching_files)


def detect_employee_column(df: pd.DataFrame) -> str | None:
    """Detect which employee-related column is available for sparkline breakdowns."""
    return next((col for col in EMPLOYEE_COLUMN_CANDIDATES if col in df.columns), None)


def apply_task_id_merges(df_curated: pd.DataFrame) -> pd.DataFrame:
    """Apply task ID merge rules and remap task/project fields to replacement task values."""
    if not TASK_ID_MERGE_MAP:
        return df_curated

    merged_df = df_curated.copy()
    task_name_by_id = (
        merged_df[['Task Id', 'Task']]
        .drop_duplicates(subset=['Task Id'])
        .set_index('Task Id')['Task']
        .to_dict()
    )
    task_project_by_id = (
        merged_df[['Task Id', 'Project Id', 'Project']]
        .drop_duplicates(subset=['Task Id'])
        .set_index('Task Id')[['Project Id', 'Project']]
        .to_dict('index')
    )

    merged_df['Task Id'] = merged_df['Task Id'].astype(int)
    for old_task_id, new_task_id in TASK_ID_MERGE_MAP.items():
        old_rows = merged_df['Task Id'] == old_task_id
        if old_rows.any():
            merged_df.loc[old_rows, 'Task Id'] = new_task_id
            replacement_name = task_name_by_id.get(new_task_id)
            if replacement_name:
                merged_df.loc[merged_df['Task Id'] == new_task_id, 'Task'] = replacement_name

            replacement_project = task_project_by_id.get(new_task_id)
            if replacement_project:
                merged_df.loc[merged_df['Task Id'] == new_task_id, 'Project Id'] = replacement_project['Project Id']
                merged_df.loc[merged_df['Task Id'] == new_task_id, 'Project'] = replacement_project['Project']
            print(f'Merged Task ID {old_task_id} into Task ID {new_task_id}')

    return merged_df


def prepare_curated_dataframe(df: pd.DataFrame, employee_col: str | None) -> pd.DataFrame:
    """Select required columns, drop invalid rows, and apply task merge rules."""
    selected_columns = REQUIRED_COLUMNS + ([employee_col] if employee_col else [])
    curated_df = df[selected_columns].copy()
    dropna_subset = REQUIRED_COLUMNS + ([employee_col] if employee_col else [])
    curated_df = curated_df.dropna(subset=dropna_subset)
    curated_df = apply_task_id_merges(curated_df)
    return curated_df


def create_sparkline_compilation_image(
    sparkline_images: list[tuple[str, Any]],
    script_dir: Path,
    latest_revision: int,
    output_prefix: str,
    file_prefix: str,
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

    compilation_file = script_dir / f'{file_prefix}_{output_prefix}_employee_sparklines_compilation_rev{latest_revision}.png'
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
    file_prefix: str,
    bar_color: str,
    script_dir: Path,
    latest_revision: int,
) -> list[tuple[str, Any]]:
    """Create one column-sparkline chart per entity id, with one row per employee."""
    if df_employee_grouped.empty:
        print("\nNo employee-level data available for sparkline plots.")
        return []

    output_images: list[tuple[str, Any]] = []

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

        # Order employees from highest to lowest total hours on this entity.
        employee_total_hours_series = (
            entity_df.groupby(entity_df[employee_col].astype(str))['Hours (h)']
            .sum()
            .sort_values(ascending=False)
        )
        employees = employee_total_hours_series.index.tolist()
        n_rows = len(employees)
        fig_height = max(2.5, (0.55 * n_rows) + 1.8) + 4.4

        fig, axes = plt.subplots(
            n_rows + 2,
            1,
            figsize=(12, fig_height),
            sharex=True,
            gridspec_kw={'height_ratios': [1.8, 1.8] + [1] * n_rows},
        )
        cumulative_ax = axes[0]
        total_ax = axes[1]
        employee_axes = axes[2:]

        total_series = (
            entity_df.groupby('MonthKey')['Hours (h)']
            .sum()
            .reindex(month_labels, fill_value=0)
        )
        cumulative_series = total_series.cumsum()
        hour_targets = PROJECT_HOUR_TARGETS.get(entity_id, {})
        sales_budget_hours = hour_targets.get('sales_budget_hours')
        allocated_hours = hour_targets.get('allocated_hours')
        configured_targets = [
            float(target)
            for target in (sales_budget_hours, allocated_hours)
            if target is not None
        ]
        cumulative_ax.plot(
            range(len(month_labels)),
            cumulative_series.values,
            color='#2a9d8f',
            marker='o',
            linewidth=2,
            markersize=4,
        )
        cumulative_ax.fill_between(
            range(len(month_labels)),
            cumulative_series.values,
            color='#2a9d8f',
            alpha=0.08,
        )
        cumulative_max_hours = max(float(cumulative_series.max()), *configured_targets, 1.0)
        cumulative_ax.set_ylim(0, cumulative_max_hours * 1.2)
        cumulative_ax.set_yticks([])
        cumulative_ax.grid(axis='y', alpha=0.2, linewidth=0.5)
        cumulative_ax.text(
            -0.08, 0.82, 'Total Team Cumulative', transform=cumulative_ax.transAxes,
            fontsize=8, fontweight='bold', va='top', ha='right', clip_on=False,
        )
        for x_pos, hour_value in enumerate(cumulative_series.values):
            if hour_value > 0:
                cumulative_ax.text(
                    x_pos,
                    float(hour_value) + (cumulative_max_hours * 0.03),
                    f"{float(hour_value):.1f}",
                    ha='center',
                    va='bottom',
                    fontsize=7,
                    fontweight='bold',
                    color='#2a9d8f',
                )
        target_line_specs = [
            ('Sales Budget Hours', sales_budget_hours, '#c1121f', '--'),
            ('Allocated Hours', allocated_hours, '#f4a261', '-.'),
        ]
        for target_label, target_hours, target_color, line_style in target_line_specs:
            if target_hours is None:
                continue
            target_value = float(target_hours)
            cumulative_ax.axhline(
                target_value,
                color=target_color,
                linestyle=line_style,
                linewidth=1.5,
            )
            cumulative_ax.text(
                0.99,
                target_value,
                f'{target_label}: {target_value:.1f}h',
                transform=cumulative_ax.get_yaxis_transform(),
                fontsize=7,
                fontweight='bold',
                color=target_color,
                va='bottom',
                ha='right',
            )
        cumulative_ax.spines['top'].set_visible(False)
        cumulative_ax.spines['right'].set_visible(False)
        cumulative_ax.spines['left'].set_visible(False)
        cumulative_ax.tick_params(axis='x', which='both', bottom=False, labelbottom=False)

        # Total team hours per month, across all employees.
        total_ax.plot(
            range(len(month_labels)),
            total_series.values,
            color='#264653',
            marker='o',
            linewidth=2,
            markersize=4,
        )
        total_ax.fill_between(range(len(month_labels)), total_series.values, color='#264653', alpha=0.08)
        total_max_hours = max(float(total_series.max()), 1.0)
        total_ax.set_ylim(0, total_max_hours * 1.2)
        total_ax.set_yticks([])
        total_ax.grid(axis='y', alpha=0.2, linewidth=0.5)
        total_ax.text(
            -0.08, 0.82, 'Total Team', transform=total_ax.transAxes,
            fontsize=8, fontweight='bold', va='top', ha='right', clip_on=False,
        )
        for x_pos, hour_value in enumerate(total_series.values):
            if hour_value > 0:
                total_ax.text(
                    x_pos,
                    float(hour_value) + (total_max_hours * 0.03),
                    f"{float(hour_value):.1f}",
                    ha='center',
                    va='bottom',
                    fontsize=7,
                    fontweight='bold',
                    color='#264653',
                )
        total_ax.spines['top'].set_visible(False)
        total_ax.spines['right'].set_visible(False)
        total_ax.spines['left'].set_visible(False)
        total_ax.tick_params(axis='x', which='both', bottom=False, labelbottom=False)

        # For visual comparison across employees on the same entity, use one shared Y-scale.
        if n_rows > 1:
            shared_entity_max_hours = max(float(entity_df['Hours (h)'].max()), 1.0)
        else:
            shared_entity_max_hours = None

        for idx, employee in enumerate(employees):
            ax = employee_axes[idx]
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

        image_label = f'{file_prefix}_{output_prefix}_{entity_id}_employee_sparklines_rev{latest_revision}'
        image_file = script_dir / f'{image_label}.png'
        fig.savefig(image_file, format='png', dpi=300, bbox_inches='tight')
        print(f"Saved sparkline image: {image_file}")

        if GENERATE_COMPILATION_IMAGE:
            buffer = io.BytesIO()
            fig.savefig(buffer, format='png', dpi=300, bbox_inches='tight')
            buffer.seek(0)
            output_images.append((image_label, mpimg.imread(buffer)))
            buffer.close()
        plt.close(fig)

    return output_images


def create_category_split_charts(
    df_project_source: pd.DataFrame,
    script_dir: Path,
    latest_revision: int,
    file_prefix: str,
) -> None:
    """Create monthly category split charts: absolute hours and 100% stacked percentage."""
    if df_project_source.empty:
        print('No project rows available for category split charts.')
        return

    category_df = df_project_source.copy()
    category_df['Project Id'] = category_df['Project Id'].astype(int)
    category_df['Category'] = category_df['Project Id'].map(PROJECT_CATEGORY_BY_ID)
    category_df = category_df[category_df['Category'].isin(CATEGORY_ORDER)].copy()

    if category_df.empty:
        print('No rows mapped to requested categories for category split charts.')
        return

    category_df['MonthKey'] = category_df['Month'].astype(str)
    grouped = (
        category_df.groupby(['MonthKey', 'Category'])['Hours (h)']
        .sum()
        .reset_index()
    )
    pivot = grouped.pivot(index='MonthKey', columns='Category', values='Hours (h)').fillna(0)
    pivot = pivot.reindex(columns=CATEGORY_ORDER, fill_value=0)

    month_order = pd.DataFrame({'month_key': pivot.index.tolist()})
    month_order['month_num'] = pd.to_numeric(month_order['month_key'], errors='coerce')
    if month_order['month_num'].notna().all():
        ordered_months = (
            month_order.sort_values(['month_num', 'month_key'])['month_key']
            .drop_duplicates()
            .tolist()
        )
    else:
        month_order['month_dt'] = pd.to_datetime(month_order['month_key'], errors='coerce')
        if month_order['month_dt'].notna().any():
            ordered_months = (
                month_order.sort_values(['month_dt', 'month_key'])['month_key']
                .drop_duplicates()
                .tolist()
            )
        else:
            ordered_months = sorted(pivot.index.tolist())

    pivot = pivot.reindex(ordered_months, fill_value=0)
    x_positions = list(range(len(pivot.index)))
    x_labels = [format_month_label(month) for month in pivot.index.tolist()]
    row_totals = pivot.sum(axis=1)
    percent_pivot = pivot.div(row_totals.replace(0, pd.NA), axis=0).fillna(0) * 100

    def save_stacked_chart(
        values_df: pd.DataFrame,
        chart_file: Path,
        title: str,
        ylabel: str,
        normalized: bool = False,
    ) -> None:
        fig, ax = plt.subplots(figsize=(12, 7))
        running_bottom = pd.Series(0.0, index=values_df.index)
        for category in CATEGORY_ORDER:
            values = values_df[category]
            hours = pivot[category]
            ax.bar(x_positions, values, bottom=running_bottom, label=category, color=CATEGORY_COLORS.get(category))
            for idx, (value, hour_value) in enumerate(zip(values, hours)):
                if value <= 0:
                    continue
                ax.text(
                    idx,
                    float(running_bottom.iloc[idx]) + float(value) / 2,
                    f'{float(hour_value):.1f}h\n{float(percent_pivot[category].iloc[idx]):.1f}%',
                    ha='center', va='center', fontsize=7, color='white', fontweight='bold',
                )
            running_bottom += values

        ax.set_xticks(x_positions)
        ax.set_xticklabels(x_labels, rotation=45, ha='right')
        ax.set_xlabel('Month')
        ax.set_ylabel(ylabel)
        if normalized:
            ax.set_ylim(0, 100)
        ax.set_title(title, fontsize=14, fontweight='bold')
        ax.legend(title='Category')
        ax.grid(axis='y', alpha=0.3)
        plt.tight_layout()
        plt.savefig(chart_file, dpi=300, bbox_inches='tight')
        plt.close(fig)

    hours_chart_file = script_dir / f'{file_prefix}_category_hours_by_month_rev{latest_revision}.png'
    save_stacked_chart(pivot, hours_chart_file, 'Monthly Hours by Category', 'Hours (h)')
    print(f'Category hours chart saved to: {hours_chart_file}')

    percent_chart_file = script_dir / f'{file_prefix}_category_percent_by_month_rev{latest_revision}.png'
    save_stacked_chart(
        percent_pivot,
        percent_chart_file,
        'Monthly Category Share (100% Stacked)',
        'Monthly Category Share (%)',
        normalized=True,
    )
    print(f'Category percent chart saved to: {percent_chart_file}')


def main() -> None:
    """Run the end-to-end sparkline workflow."""
    global PROJECT_CATEGORY_BY_ID, PROJECT_REFERENCE_BY_ID

    script_dir = Path(__file__).parent
    execution_time = datetime.now().astimezone()
    executed_at = execution_time.strftime('%Y-%m-%d %H:%M:%S %Z')
    file_prefix = f'Defense_LP_Project_{execution_time.strftime("%Y-%m-%d_%H-%M-%S")}'
    latest_revision, latest_file, file_count = find_latest_timesheet_export(script_dir)

    print(f'Found {file_count} timesheet_export file(s)')
    print(f'Opening latest revision: {latest_file.name} (revision {latest_revision})')

    df = pd.read_excel(latest_file, header=0)

    print(f'\nLoaded {len(df)} rows and {len(df.columns)} columns')
    print(f'\nColumns: {list(df.columns)}')
    print('\nFirst few rows:')
    print(df.head())

    print('\nDataFrame info:')
    df.info()

    employee_col = detect_employee_column(df)
    if employee_col is None:
        print('\nWarning: No employee column found. Sparkline-by-employee plots will be skipped.')
    else:
        print(f'\nUsing employee column for sparklines: {employee_col}')

    missing_columns = [col for col in REQUIRED_COLUMNS + [EMPLOYEE_NUMBER_COLUMN] if col not in df.columns]
    if missing_columns:
        print(f'\nWarning: Missing columns: {missing_columns}')
        print('Available columns:', list(df.columns))
        return

    defense_employee_numbers = set(DEFENSE_EMPLOYEE_NUMBER_BY_LP_ID.values())
    defense_employee_rows = df[
        df[EMPLOYEE_NUMBER_COLUMN].isin(defense_employee_numbers)
        & (df['Hours (h)'] > 0)
    ].copy()
    discovered_project_ids = sorted(
        defense_employee_rows['Project Id'].dropna().astype(int).unique().tolist()
    )
    if not discovered_project_ids:
        print('\nNo projects found for the configured Defense employees.')
        return

    print(
        f'\nDiscovered {len(discovered_project_ids)} projects from '
        f'{defense_employee_rows[EMPLOYEE_NUMBER_COLUMN].nunique()} Defense employees.'
    )

    df_curated = prepare_curated_dataframe(defense_employee_rows, employee_col)
    df_project_source = df_curated[df_curated['Project Id'].isin(discovered_project_ids)].copy()
    PROJECT_CATEGORY_BY_ID = {
        project_id: (
            'ADMIN' if project_id in ADMIN_PROJECT_IDS
            else 'SUPPORT' if project_id in SUPPORT_PROJECT_IDS
            else 'NRE'
        )
        for project_id in discovered_project_ids
    }
    PROJECT_REFERENCE_BY_ID = (
        df_project_source[['Project Id', 'Project']]
        .drop_duplicates(subset=['Project Id'])
        .assign(**{'Project Id': lambda data: data['Project Id'].astype(int)})
        .set_index('Project Id')['Project']
        .astype(str)
        .to_dict()
    )

    df_project_grouped = (
        df_project_source.groupby(['Month', 'Project Id', 'Project'])['Hours (h)']
        .sum()
        .reset_index()
    )
    project_total_hours_by_id = (
        df_project_source.groupby('Project Id')['Hours (h)']
        .sum()
        .astype(float)
        .to_dict()
    )
    report_project_ids = []
    for category in CATEGORY_ORDER:
        category_project_ids = [
            project_id for project_id in discovered_project_ids
            if (
                PROJECT_CATEGORY_BY_ID.get(project_id) == category
                and project_total_hours_by_id.get(project_id, 0.0) >= MIN_REPORT_PROJECT_HOURS
            )
        ]
        category_project_ids.sort(
            key=lambda project_id: project_total_hours_by_id.get(project_id, 0.0),
            reverse=True,
        )
        report_project_ids.extend(category_project_ids)
    project_month_hours_by_id = {}
    for project_id, project_df in df_project_source.groupby('Project Id'):
        monthly_totals = project_df.groupby('Month')['Hours (h)'].sum()
        project_month_hours_by_id[int(project_id)] = {
            format_month_label(str(month)): float(hours) for month, hours in monthly_totals.items()
        }
    project_month_labels = [
        format_month_label(str(month)) for month in sorted(df_project_source['Month'].dropna().unique())
    ]
    df_project_filtered = df_project_grouped.copy()

    print('\nFiltered data summary (specific Project Ids):')
    print(f'Raw rows matched on Project Id column: {len(df_project_source)}')
    print(f"Total projects included: {df_project_filtered['Project'].nunique()}")
    print(f"Total months: {df_project_filtered['Month'].nunique()}")
    print('\nFiltered Project data preview:')
    print(df_project_filtered.head(10))
    missing_project_ids = sorted(set(discovered_project_ids) - set(df_project_source['Project Id'].dropna().astype(int).tolist()))
    if missing_project_ids:
        print(f'Warning: Requested Project Ids not found in grouped data: {missing_project_ids}')

    create_category_split_charts(
        df_project_source=df_project_source,
        script_dir=script_dir,
        latest_revision=latest_revision,
        file_prefix=file_prefix,
    )

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
            entity_ids=report_project_ids,
            title_prefix_col='Project',
            title_prefix_label='Project',
            title_entity_label='Project',
            output_prefix='project',
            file_prefix=file_prefix,
            bar_color='#1d4ed8',
            script_dir=script_dir,
            latest_revision=latest_revision,
        )
        if GENERATE_COMPILATION_IMAGE:
            create_sparkline_compilation_image(
                sparkline_images=project_sparkline_images,
                script_dir=script_dir,
                latest_revision=latest_revision,
                output_prefix='all_project',
                file_prefix=file_prefix,
                compilation_title='Employee Sparkline Compilation - All Projects',
            )
        cleanup_stale_entity_images(
            script_dir=script_dir,
            latest_revision=latest_revision,
            keep_entity_ids=report_project_ids,
            output_prefix='project',
            file_prefix=file_prefix,
        )

    # When selected, break the rollup project into one chart per Task Id,
    # ordered from the highest total hours to the lowest.
    df_task_breakdown_source = df_project_source[df_project_source['Project Id'] == TASK_BREAKDOWN_PROJECT_ID].copy()
    task_hours_totals = df_task_breakdown_source.groupby('Task Id')['Hours (h)'].sum().sort_values(ascending=False)
    task_ids = task_hours_totals.index.dropna().astype(int).tolist()
    task_total_hours_by_id = {
        int(task_id): float(total_hours)
        for task_id, total_hours in task_hours_totals.items()
    }
    task_name_by_id: dict[int, str] = {}
    if task_ids:
        df_task_grouped = (
            df_task_breakdown_source.groupby(['Month', 'Task Id', 'Task', 'Project Id', 'Project'])['Hours (h)']
            .sum()
            .reset_index()
        )
        task_output_csv = script_dir / f'{file_prefix}_task_hours_by_month_data_rev{latest_revision}.csv'
        df_task_grouped.to_csv(task_output_csv, index=False)
        print(f'Task breakdown data saved to: {task_output_csv}')

        task_name_by_id = (
            df_task_breakdown_source[['Task Id', 'Task']]
            .dropna()
            .drop_duplicates(subset=['Task Id'])
            .assign(**{'Task Id': lambda d: d['Task Id'].astype(int)})
            .set_index('Task Id')['Task']
            .to_dict()
        )

        if employee_col:
            df_task_employee_grouped = (
                df_task_breakdown_source.groupby(['Month', 'Task Id', 'Task', 'Project', employee_col])['Hours (h)']
                .sum()
                .reset_index()
            )
            task_sparkline_images = create_entity_employee_sparkline_plots(
                df_employee_grouped=df_task_employee_grouped,
                employee_col=employee_col,
                entity_id_col='Task Id',
                entity_name_col='Task',
                entity_ids=task_ids,
                title_prefix_col='Project',
                title_prefix_label='Project',
                title_entity_label='Task',
                output_prefix='task',
                file_prefix=file_prefix,
                bar_color=TASK_BAR_COLOR,
                script_dir=script_dir,
                latest_revision=latest_revision,
            )
            if GENERATE_COMPILATION_IMAGE:
                create_sparkline_compilation_image(
                    sparkline_images=task_sparkline_images,
                    script_dir=script_dir,
                    latest_revision=latest_revision,
                    output_prefix='all_task',
                    file_prefix=file_prefix,
                    compilation_title=f'Employee Sparkline Compilation - Tasks in Project {TASK_BREAKDOWN_PROJECT_ID}',
                )
            cleanup_stale_entity_images(
                script_dir=script_dir,
                latest_revision=latest_revision,
                keep_entity_ids=task_ids,
                output_prefix='task',
                file_prefix=file_prefix,
            )

    if GENERATE_VERTICAL_BAR_CHART:
        df_pivot = df_project_filtered.pivot_table(
            index='Month',
            columns=['Project Id', 'Project'],
            values='Hours (h)',
            aggfunc='sum',
        )
        fig, ax = plt.subplots(figsize=(12, 8))
        df_pivot.plot(kind='bar', stacked=True, ax=ax)

        x_labels = [format_month_label(str(idx)) for idx in df_pivot.index]
        ax.set_xticklabels(x_labels, rotation=45, ha='right')
        ax.set_xlabel('Month', fontsize=12)
        ax.set_ylabel('Hours (h)', fontsize=12)
        ax.set_title('Hours Spent by Month on Specific Projects', fontsize=14, fontweight='bold')
        ax.legend(title='Project', bbox_to_anchor=(1.05, 1), loc='upper left')
        ax.grid(True, alpha=0.3)

        plt.tight_layout()
        output_file = script_dir / f'{file_prefix}_hours_by_month_plot_rev{latest_revision}.png'
        plt.savefig(output_file, dpi=300, bbox_inches='tight')
        print(f'\nPlot saved to: {output_file}')
        plt.close(fig)
    else:
        print('\nVertical bar chart generation is disabled.')

    output_csv = script_dir / f'{file_prefix}_hours_by_month_project_data_rev{latest_revision}.csv'
    df_project_filtered.to_csv(output_csv, index=False)
    print(f'Curated data saved to: {output_csv}')

    if AUTO_RENDER_QMD_REPORT:
        render_succeeded = render_qmd_report(script_dir, file_prefix, executed_at)
        pdf_succeeded = generate_project_pdf_report(
            script_dir=script_dir,
            latest_revision=latest_revision,
            output_prefix=file_prefix,
            executed_at=executed_at,
            ordered_project_ids=report_project_ids,
            task_breakdown_project_id=TASK_BREAKDOWN_PROJECT_ID,
            ordered_task_ids=task_ids,
            task_name_by_id=task_name_by_id,
            project_total_hours_by_id=project_total_hours_by_id,
            task_total_hours_by_id=task_total_hours_by_id,
            project_month_hours_by_id=project_month_hours_by_id,
            project_month_labels=project_month_labels,
        )
        if render_succeeded and pdf_succeeded:
            delete_all_png_files(script_dir)


if __name__ == '__main__':
    main()
