"""Run the PSE, Defense, and CleanEnergy time-allocation reports."""

import os
from pathlib import Path
import re
import subprocess
import sys
import tempfile
import warnings

import pandas as pd


SCRIPTS = (
    'PSE_LP_Project-plotting.py',
    'Defense_LP_Project-plotting.py',
    'CleanEnergy_LP_Project-plotting.py',
)

QUARTERLY_INPUTS = (
    'timesheet_export_Q126.xlsx',
    'timesheet_export_Q226.xlsx',
    'timesheet_export_Q326.xlsx',
    'timesheet_export_Q426.xlsx',
)
TIMESHEET_COLUMNS = (
    'Date',
    'For',
    'Hours (h)',
    'Task Id',
    'Task',
    'Project Id',
    'Project',
    'Month',
    'Employee Number (Profile)',
)
TIMESHEET_EXPORT_PATH_ENV = 'LP_TIMESHEET_EXPORT_PATH'


def load_combined_timesheet(script_dir: Path) -> pd.DataFrame:
    """Combine the quarter exports, taking only Q4-period rows from the YTD file."""
    quarter_frames = []
    quarter_pattern = re.compile(r'timesheet_export_Q([1-4])(\d{2})\.xlsx$', re.IGNORECASE)

    for input_name in QUARTERLY_INPUTS:
        input_path = script_dir / input_name
        if not input_path.is_file():
            raise FileNotFoundError(f'Required quarterly timesheet export not found: {input_path}')

        with warnings.catch_warnings():
            warnings.filterwarnings(
                'ignore',
                message="Workbook contains no default style, apply openpyxl's default",
                category=UserWarning,
                module=r'openpyxl\.styles\.stylesheet',
            )
            frame = pd.read_excel(
                input_path,
                header=0,
                usecols=lambda column: column in TIMESHEET_COLUMNS,
            )
        missing_columns = [column for column in TIMESHEET_COLUMNS if column not in frame.columns]
        if missing_columns:
            raise ValueError(f'{input_path.name} is missing required columns: {missing_columns}')

        frame['Date'] = pd.to_datetime(frame['Date'], errors='coerce')
        if frame['Date'].isna().any():
            invalid_rows = int(frame['Date'].isna().sum())
            raise ValueError(f'{input_path.name} contains {invalid_rows} rows with missing or invalid dates.')

        match = quarter_pattern.fullmatch(input_name)
        if match is None:
            raise ValueError(f'Unexpected quarterly export filename: {input_name}')
        quarter_number = int(match.group(1))
        year = 2000 + int(match.group(2))
        quarter = pd.Period(f'{year}Q{quarter_number}', freq='Q-DEC')

        if quarter_number == 4:
            frame = frame[
                frame['Date'].ge(quarter.start_time)
                & frame['Date'].lt((quarter + 1).start_time)
            ].copy()
        elif not frame['Date'].dt.to_period('Q-DEC').eq(quarter).all():
            raise ValueError(f'{input_path.name} contains rows outside its named fiscal quarter.')

        quarter_frames.append(frame.loc[:, TIMESHEET_COLUMNS])

    return pd.concat(quarter_frames, ignore_index=True)


def main() -> int:
    script_dir = Path(__file__).resolve().parent
    failures = []

    combined_timesheet = load_combined_timesheet(script_dir)
    print(f'Combined {len(combined_timesheet)} timesheet rows from the quarterly exports.')

    export_pattern = re.compile(r'^timesheet_export \((\d+)\)\.xlsx$', re.IGNORECASE)
    revisions = [
        int(match.group(1))
        for path in script_dir.glob('timesheet_export*.xlsx')
        if (match := export_pattern.fullmatch(path.name))
    ]
    next_revision = max(revisions, default=0) + 1

    with tempfile.TemporaryDirectory(prefix='dptech-timesheet-') as temporary_dir:
        combined_path = Path(temporary_dir) / f'timesheet_export ({next_revision}).xlsx'
        combined_timesheet.to_excel(combined_path, index=False, sheet_name='Timesheets')
        report_env = os.environ.copy()
        report_env[TIMESHEET_EXPORT_PATH_ENV] = str(combined_path)

        for script_name in SCRIPTS:
            script_path = script_dir / script_name
            print(f'\n=== Running {script_name} ===', flush=True)

            if not script_path.is_file():
                print(f'Script not found: {script_path}')
                failures.append(script_name)
                continue

            result = subprocess.run(
                [sys.executable, str(script_path)],
                cwd=script_dir,
                env=report_env,
                check=False,
            )
            if result.returncode != 0:
                failures.append(f'{script_name} (exit code {result.returncode})')

    if failures:
        print('\nBatch completed with failures:')
        for failure in failures:
            print(f'- {failure}')
        return 1

    print('\nAll report workflows completed successfully.')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())