#!/usr/bin/env python3
import pygetwindow as gw
from prettytable import PrettyTable
from datetime import datetime
from object import event
import param
import config
import time
import signal
import sys
import json
import xlsxwriter
import os.path

"""
Init params
"""
events = param.data_init()
windows = param.data_init()
j_file = param.file_json
e_file = param.file_excel


def containsWindow(w, t):
	for e in w:
		if str(e) == str(t):
			return True
	return False


def addEvent(item, date):
	text = item.split(" - ")[-1]
	text = text.split(" | ")[-1]

	if containsWindow(windows, text):
		idx = windows.index(text)
		events[idx].count = events[idx].count + 1
		events[idx].last = date
	else:
		windows.append(text)
		evt = event(text, date)
		events.append(evt)


def percentage(total, value):
	return round((value * 100) / total, 2)


def elapsedTime(start, end):
	diff = end - start
	print(f">>> Elapsed time: {GetTimer(diff)}\n")
	return diff


def GetTimer(sec):
	hours, rem = divmod(sec, 3600)
	minutes, seconds = divmod(rem, 60)
	return "{:0>2}:{:0>2}:{:05.4f}".format(int(hours), int(minutes), seconds)


def getDuration(value):
	count = 0.90 / config.interval
	seconds = value / count
	return GetTimer(seconds)


def status():
	end = time.time()
	full = 0
	for e in events:
		full = full + e.count
	events_table = PrettyTable(["FIRST", "LAST", "DURATION", "%", "TITLE"])
	sorted_events = sorted(events, key=lambda x: x.count, reverse=True)

	# Excel report
	workbook = xlsxwriter.Workbook(e_file)
	worksheet = workbook.add_worksheet("Data")
	bold = workbook.add_format({"bold": True})
	number_format = workbook.add_format({"num_format": "##0.00"})
	worksheet.write("A1", "FIRST", bold)
	worksheet.write("B1", "LAST", bold)
	worksheet.write("C1", "DURATION", bold)
	worksheet.write("D1", "%", bold)
	worksheet.write("E1", "TITLE", bold)

	row = 1
	col = 0
	total_p = 0.0

	for e in sorted_events:
		p = percentage(full, e.count)
		w = e.title[-config.title_size :]
		d = getDuration(e.count)
		if p > config.ignore_event:
			events_table.add_row([e.first, e.last, d, p, w])
			worksheet.write(row, col, e.first)
			worksheet.write(row, col + 1, e.last)
			worksheet.write_string(row, col + 2, d)
			worksheet.write_number(row, col + 3, p, number_format)
			worksheet.write_string(row, col + 4, w)
			row += 1
			total_p += p

	if total_p < 100.0:
		others = 100.0 - total_p
		worksheet.write(row, col, "N/A")
		worksheet.write(row, col + 1, "N/A")
		worksheet.write(row, col + 2, "N/A")
		worksheet.write_number(row, col + 3, others, number_format)
		worksheet.write_string(row, col + 4, "Others")
		row += 1

	pie_chart = workbook.add_chart({"type": "pie"})
	pie_chart.set_legend({"none": True})
	pie_chart.add_series(
		{
			"name": "My productivity today",
			"categories": f"=Data!$E$2:$E${row}",
			"values": f"=Data!$D$2:$D${row}",
			"data_labels": {
				"value": True,
				"category_name": True,
				"position": "outside_end",
			},
		}
	)
	pie_chart.set_style(10)
	pie_chart.set_legend({"position": "right"})
	pie_chart.set_size({"width": 720, "height": 560})
	worksheet.insert_chart("G3", pie_chart)

	worksheet.write(f"A{row + 2}", "ELAPSED", bold)
	worksheet.write(f"A{row + 3}", GetTimer(end - start))
	workbook.close()

	# Output
	sys.stdout.write("\r" + str(events_table) + "\n")
	sys.stdout.flush()
	diff = elapsedTime(start, end)

	# JSON Dump
	with open(j_file, mode="w", encoding="utf-8") as outfile:
		data = [obj.__dict__ for obj in sorted_events]
		out = {"elapsed": diff, "data": data}
		json.dump(out, outfile)


def signalHandler(signal, frame):
	print("\033c")
	print(f"SUMMARY\n")
	status()
	sys.exit(0)


def animation(value):
	animation = "|/-\\"
	sys.stdout.write("\r" + animation[value % len(animation)])
	sys.stdout.flush()


"""
Main
"""
if __name__ == "__main__":
	signal.signal(signal.SIGINT, signalHandler)
	print("[INFO] Start monitoring, press CTRL + C to stop")
	progress = 0
	time.sleep(1)
	start = time.time()

	# Check Dump
	if os.path.exists(j_file):
		print("[INFO] Load existing data")

		with open(j_file, "r") as read_file:
			data = json.load(read_file)

		start = start - data.get("elapsed")

		for e in data.get("data"):
			windows.append(e.get("title"))
			evt = event(e.get("title"), e.get("first"))
			evt.last = e.get("last")
			evt.count = e.get("count")
			events.append(evt)

	while True:
		try:
			win = gw.getActiveWindow().title
			dt = datetime.now().strftime("%d/%m/%y %H:%M:%S")
			opened = len(gw.getAllTitles())
			if win:
				addEvent(win, dt)
				if config.debug:
					total = len(events)
					print(f"[DEBUG] {dt} | {opened} | {total} | {win}")
				else:
					animation(progress)
		except Exception as err:
			if config.debug:
				print(f"[ERROR] {str(err)}")
			pass

		time.sleep(config.interval)
		progress += 1
		if progress > config.checkpoint:
			progress = 0
			status()
