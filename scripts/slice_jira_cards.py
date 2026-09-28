#!/usr/bin/env python3
"""
Generate src/jira-cards-current.csv: a current-month slice of jira-cards.csv.

Used by the TV browser rendering path to avoid loading the full 20+ MB file.
The sliced file contains only rows whose Created date matches the current
calendar month (YYYY-MM), reducing download size by ~97%.

Run automatically by update-service-catalog.yml after export_itsm_cards.py.
Can also be run manually:  python scripts/slice_jira_cards.py
"""
import csv
import datetime
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(SCRIPT_DIR, '..', 'src', 'jira-cards.csv')
DST = os.path.join(SCRIPT_DIR, '..', 'src', 'jira-cards-current.csv')

CREATED_COL = 3  # zero-based index of the 'Created' column


def main():
    today = datetime.date.today()
    prefix = today.strftime('%Y-%m')

    if not os.path.exists(SRC):
        print(f'[ERROR] Source not found: {SRC}', file=sys.stderr)
        sys.exit(1)

    written = 0
    with open(SRC, newline='', encoding='utf-8') as f_in, \
         open(DST, 'w', newline='', encoding='utf-8') as f_out:

        f_out.write(f'# generated: {today.isoformat()}T00:00:00Z (current-month slice: {prefix})\n')

        reader = csv.reader(f_in)
        writer = csv.writer(f_out)

        for row in reader:
            if not row:
                continue
            if row[0].startswith('#'):
                continue
            if row[0] == 'Key':
                writer.writerow(row)
                continue
            if len(row) > CREATED_COL and row[CREATED_COL].startswith(prefix):
                writer.writerow(row)
                written += 1

    print(f'[slice_jira_cards] {prefix}: {written} rows → {DST}')


if __name__ == '__main__':
    main()
