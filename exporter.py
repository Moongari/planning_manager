# exporter.py
import xlsxwriter
from datetime import timedelta
import calendar

def export_to_excel(persons, task_manager, repos, start_date, filename):
    wb = xlsxwriter.Workbook(filename)
    ws = wb.add_worksheet("Planning")

    bold = wb.add_format({"bold": True})
    wrap = wb.add_format({"text_wrap": True})
    percent_fmt = wb.add_format({"num_format": "0%"})
    orange = wb.add_format({"bg_color": "#FFA500"})

    days = [(start_date + timedelta(days=i)) for i in range(7)]
    ws.write(0, 0, "Personne", bold)
    for i, d in enumerate(days):
        ws.write(0, i+1, d.strftime("%A %d/%m"), bold)

    for r, person in enumerate(persons, start=1):
        ws.write(r, 0, person, bold)
        for c, day in enumerate(days):
            day_name = calendar.day_name[day.weekday()]
            is_repos = day_name in repos[person]
            cell_text = "Repos" if is_repos else ""

            if not is_repos:
                tasks = task_manager.get_tasks_for(person, day.strftime("%Y-%m-%d"))
                total_duration = sum(t["duration"] for t in tasks)
                content = "\n".join(f"{t['name']} ({t['duration']}m)" for t in tasks)
                percent = total_duration / (8 * 60)
                cell_text = f"{content}\n{int(percent * 100)}%"

            fmt = orange if is_repos else wrap
            ws.write(r, c + 1, cell_text, fmt)

    wb.close()
