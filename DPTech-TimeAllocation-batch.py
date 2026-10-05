"""Run the PSE, Defense, and CleanEnergy time-allocation reports."""

from pathlib import Path
import subprocess
import sys


SCRIPTS = (
    'PSE_LP_Project-plotting.py',
    'Defense_LP_Project-plotting.py',
    'CleanEnergy_LP_Project-plotting.py',
)


def main() -> int:
    script_dir = Path(__file__).resolve().parent
    failures = []

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