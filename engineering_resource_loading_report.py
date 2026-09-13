"""Prototype engineering resource loading report."""


from datetime import date, timedelta
from collections import defaultdict
from pathlib import Path

import pandas as pd


ENGINEERING_START_DELAY_DAYS = 14
numberEngineers = 2
avgAppliedHoursPerDay = 6


def _add_working_days(start_date, working_days):
	current_date = start_date
	days_added = 0
	while days_added < working_days:
		current_date += timedelta(days=1)
		if current_date.weekday() < 5:
			days_added += 1
	return current_date


def _build_daily_engineering_schedule(start_date, dwg_release, sched_date, applied_hours_per_day):
	daily_schedule = []
	current_date = start_date
	while current_date <= sched_date:
		if current_date.weekday() < 5:
			load_percent = 1.0 if current_date <= dwg_release else 0.15
			daily_schedule.append(
				{
					"date": current_date,
					"load_percent": load_percent,
					"applied_hours": applied_hours_per_day * load_percent,
				}
			)
		current_date += timedelta(days=1)
	return daily_schedule


def calculate_schedule_fields(customer_project):
	ord_date = date.fromisoformat(customer_project["ORD_DATE"])
	sched_date = date.fromisoformat(customer_project["Sched_Date"])
	if sched_date < ord_date:
		raise ValueError("Sched_Date must be on or after ORD_DATE")

	duration = sum(
		1
		for day_offset in range((sched_date - ord_date).days)
		if (ord_date + timedelta(days=day_offset + 1)).weekday() < 5
	)
	applied_hours_per_day = (
		customer_project["TRUE_EngrHoursBudget"] / duration / 1.15
	)
	engineering_start_date = ord_date + timedelta(days=ENGINEERING_START_DELAY_DAYS)
	bom_release = _add_working_days(ord_date, round(duration / 3))
	dwg_release = _add_working_days(ord_date, round(duration * 2 / 3))
	return {
		**customer_project,
		"duration": duration,
		"appliedHoursPerDay": applied_hours_per_day,
		"Engineering_Start_Date": engineering_start_date,
		"BOM_release": bom_release,
		"DWG_release": dwg_release,
		"daily_engineering_schedule": _build_daily_engineering_schedule(
			engineering_start_date,
			dwg_release,
			sched_date,
			applied_hours_per_day,
		),
	}


def build_total_daily_applied_hours(project_schedules):
	daily_totals = defaultdict(float)
	daily_project_counts = defaultdict(int)
	for project in project_schedules:
		for daily_entry in project["daily_engineering_schedule"]:
			current_date = daily_entry["date"]
			daily_totals[current_date] += daily_entry["applied_hours"]
			daily_project_counts[current_date] += 1

	return [
		{
			"date": current_date,
			"appliedHoursPerDay": daily_totals[current_date],
			"active_project_count": daily_project_counts[current_date],
		}
		for current_date in sorted(daily_totals)
	]


def build_daily_net_capacity(total_daily_applied_hours):
	if not total_daily_applied_hours:
		return []

	capacity_per_day = numberEngineers * avgAppliedHoursPerDay
	daily_applied_hours = {
		row["date"]: row["appliedHoursPerDay"]
		for row in total_daily_applied_hours
	}
	first_date = min(daily_applied_hours)
	last_date = max(daily_applied_hours)
	net_capacity = []
	current_date = first_date
	while current_date <= last_date:
		if current_date.weekday() < 5:
			applied_hours = daily_applied_hours.get(current_date, 0.0)
			net_capacity.append(
				{
					"date": current_date,
					"appliedHoursPerDay": applied_hours,
					"capacityHoursPerDay": capacity_per_day,
					"net_capacity": applied_hours - capacity_per_day,
				}
			)
		current_date += timedelta(days=1)
	return net_capacity


def plot_daily_net_capacity(net_capacity, output_path):
	import matplotlib.pyplot as plt

	dates = [row["date"] for row in net_capacity]
	net_values = [row["net_capacity"] for row in net_capacity]
	over_values = [value if value > 0 else None for value in net_values]
	under_values = [value if value < 0 else None for value in net_values]

	figure, axis = plt.subplots(figsize=(16, 7))
	axis.plot(dates, over_values, color="#c0392b", linewidth=2, label="Over-applied")
	axis.plot(dates, under_values, color="#2e8b57", linewidth=2, label="Under-applied")
	axis.axhline(0, color="black", linewidth=1)
	axis.set_title("Daily Engineering Net Capacity")
	axis.set_xlabel("Date")
	axis.set_ylabel("Applied hours minus available capacity")
	axis.legend()
	axis.grid(True, axis="y", alpha=0.3)
	figure.autofmt_xdate()
	figure.tight_layout()

	output_path = Path(output_path)
	output_path.parent.mkdir(parents=True, exist_ok=True)
	figure.savefig(output_path, dpi=150)
	plt.close(figure)
	return output_path


def build_monthly_project_applied_hours(project_schedules):
	monthly_totals = defaultdict(float)
	for project in project_schedules:
		for daily_entry in project["daily_engineering_schedule"]:
			month = daily_entry["date"].replace(day=1)
			monthly_totals[(month, project["customer"])] += daily_entry["applied_hours"]

	return [
		{
			"month": month,
			"customer": customer,
			"applied_hours": monthly_totals[(month, customer)],
		}
		for month, customer in sorted(monthly_totals)
	]


def plot_monthly_project_workload(monthly_workload, output_path):
	import matplotlib.pyplot as plt

	months = sorted({row["month"] for row in monthly_workload})
	customers = sorted({row["customer"] for row in monthly_workload})
	month_labels = [month.strftime("%b %Y") for month in months]
	bottom = [0.0] * len(months)

	figure, axis = plt.subplots(figsize=(14, 7))
	for customer in customers:
		values = [
			next(
				row["applied_hours"]
				for row in monthly_workload
				if row["month"] == month and row["customer"] == customer
			)
			if any(
				row["month"] == month and row["customer"] == customer
				for row in monthly_workload
			)
			else 0.0
			for month in months
		]
		axis.bar(month_labels, values, bottom=bottom, label=customer)
		bottom = [current + value for current, value in zip(bottom, values)]

	axis.set_title("Monthly Engineering Workload by Customer")
	axis.set_xlabel("Month")
	axis.set_ylabel("Applied engineering hours")
	axis.tick_params(axis="x", rotation=45)
	axis.legend(title="Customer")
	figure.tight_layout()

	output_path = Path(output_path)
	output_path.parent.mkdir(parents=True, exist_ok=True)
	figure.savefig(output_path, dpi=150)
	plt.close(figure)
	return output_path


def build_iso_week_project_applied_hours(project_schedules):
	weekly_totals = defaultdict(float)
	for project in project_schedules:
		for daily_entry in project["daily_engineering_schedule"]:
			iso_year, iso_week, _ = daily_entry["date"].isocalendar()
			week_label = f"{iso_year}-W{iso_week:02d}"
			weekly_totals[(week_label, project["customer"])] += daily_entry["applied_hours"]

	return [
		{
			"week": week_label,
			"customer": customer,
			"applied_hours": weekly_totals[(week_label, customer)],
		}
		for week_label, customer in sorted(weekly_totals)
	]


def plot_fiscal_week_project_workload(project_schedules, output_path):
	import matplotlib.pyplot as plt

	weekly_workload = build_iso_week_project_applied_hours(project_schedules)
	weeks = sorted({row["week"] for row in weekly_workload})
	customer_totals = {
		customer: sum(
			row["applied_hours"]
			for row in weekly_workload
			if row["customer"] == customer
		)
		for customer in {row["customer"] for row in weekly_workload}
	}
	def shorten_legend_label(project_name, sodet_key):
		base_label = f"{sodet_key} - {project_name}"
		if len(base_label) <= 28:
			return base_label
		trimmed_name = project_name[:18].rstrip()
		return f"{sodet_key} - {trimmed_name}..."

	customer_display_names = {
		project["customer"]: shorten_legend_label(project["customer"], project["SODET_KEY"])
		for project in project_schedules
	}
	customers = sorted(customer_totals, key=lambda customer: customer_totals[customer], reverse=True)
	bottom = {week_label: 0.0 for week_label in weeks}
	week_positions = list(range(len(weeks)))
	month_boundary_positions = []
	month_interval_labels = []
	last_month_label = None
	last_month_start_index = 0
	for week_index, week_label in enumerate(weeks):
		iso_year_text, iso_week_text = week_label.split("-W")
		iso_year = int(iso_year_text)
		iso_week = int(iso_week_text)
		week_start = date.fromisocalendar(iso_year, iso_week, 1)
		month_label = week_start.strftime("%b %Y")
		if last_month_label is None:
			last_month_label = month_label
			last_month_start_index = week_index
		elif month_label != last_month_label:
			month_boundary_positions.append(week_index - 0.5)
			month_interval_labels.append(((last_month_start_index + week_index - 1) / 2, last_month_label))
			last_month_label = month_label
			last_month_start_index = week_index
	if last_month_label is not None:
		month_interval_labels.append(((last_month_start_index + len(weeks) - 1) / 2, last_month_label))

	figure, axis = plt.subplots(figsize=(20, 8))
	for customer in customers:
		values = [
			next(
				row["applied_hours"]
				for row in weekly_workload
				if row["week"] == week_label and row["customer"] == customer
			)
			if any(
				row["week"] == week_label and row["customer"] == customer
				for row in weekly_workload
			)
			else 0.0
			for week_label in weeks
		]
		axis.bar(
			week_positions,
			values,
			bottom=[bottom[week_label] for week_label in weeks],
			label=customer_display_names.get(customer, customer),
			width=0.8,
			alpha=0.8,
		)
		for week_label in weeks:
			bottom[week_label] += next(
				(row["applied_hours"] for row in weekly_workload if row["week"] == week_label and row["customer"] == customer),
				0.0,
			)

	for boundary in month_boundary_positions:
		axis.axvline(boundary, color="#b0b0b0", linewidth=1.2, alpha=0.9)
	for label_x, label in month_interval_labels:
		axis.text(
			label_x,
			axis.get_ylim()[1] * 0.98,
			label,
			rotation=0,
			ha="center",
			va="bottom",
			fontsize=7,
			color="#666666",
		)

	axis.set_title("Engineering Workload by ISO Fiscal Week and Customer", fontsize=14)
	axis.set_xlabel("ISO Fiscal Week", fontsize=11)
	axis.set_ylabel("Applied engineering hours", fontsize=11)
	axis.grid(True, axis="y", alpha=0.3)
	axis.tick_params(axis="x", rotation=90, labelsize=8)
	axis.tick_params(axis="y", labelsize=9)
	axis.set_xticks(week_positions)
	axis.set_xticklabels(weeks)
	axis.margins(x=0.01)
	axis.legend(
		title="Customer",
		loc="upper right",
		frameon=False,
		fontsize=8,
		title_fontsize=9,
	)
	figure.subplots_adjust(left=0.07, right=0.98, bottom=0.2, top=0.9)

	output_path = Path(output_path)
	if not output_path.is_absolute():
		output_path = Path.cwd() / output_path
	output_path.parent.mkdir(parents=True, exist_ok=True)
	figure.savefig(output_path, dpi=150)
	plt.close(figure)
	return output_path


REAL_PROJECTS_PATH = Path("C:/Users/mwoodmansee/OneDrive - Dynapower Company/OneDrive - Operations 26/PSE-project_engrHours.xlsx")


def load_customer_projects(excel_path):
	if not excel_path.exists():
		raise FileNotFoundError(f"Workbook not found: {excel_path}")

	df = pd.read_excel(excel_path)
	df.columns = [str(column).strip() for column in df.columns]

	projects = []
	for row in df.to_dict(orient="records"):
		job_name = row.get("Job Name")
		sodet_key = row.get("Sodet Key")
		sales_budget = pd.to_numeric(row.get("Sales Budget (hrs)"), errors="coerce")
		ord_date = pd.to_datetime(row.get("ORD DATE"), errors="coerce")
		sched_date = pd.to_datetime(row.get("Sched Date"), errors="coerce")

		if pd.isna(sales_budget) or float(sales_budget) <= 10:
			continue
		if pd.isna(ord_date) or pd.isna(sched_date):
			continue
		if pd.isna(job_name) or pd.isna(sodet_key):
			continue

		projects.append(
			{
				"customer": str(job_name).strip(),
				"SODET_KEY": str(sodet_key).strip(),
				"ORD_DATE": ord_date.date().isoformat(),
				"Sched_Date": sched_date.date().isoformat(),
				"Allocated_hrs": row.get("Allocated (hrs)"),
				"Actual_hrs": row.get("Actual (hrs)"),
				"TRUE_EngrHoursBudget": float(sales_budget),
				"GrossPrice": float(pd.to_numeric(row.get("Gross Price"), errors="coerce") or 0),
				"available_hours": float(sales_budget),
			}
		)
	return projects


CUSTOMER_PROJECTS = load_customer_projects(REAL_PROJECTS_PATH)

PROJECT_SCHEDULES = [calculate_schedule_fields(project) for project in CUSTOMER_PROJECTS]
TOTAL_DAILY_APPLIED_HOURS = build_total_daily_applied_hours(PROJECT_SCHEDULES)
DAILY_NET_CAPACITY = build_daily_net_capacity(TOTAL_DAILY_APPLIED_HOURS)
MONTHLY_PROJECT_APPLIED_HOURS = build_monthly_project_applied_hours(PROJECT_SCHEDULES)


if __name__ == "__main__":
	plot_fiscal_week_project_workload(
		PROJECT_SCHEDULES,
		Path("engineering_resource_loading_by_fiscal_week.png"),
	)
