#!/usr/bin/env python3
import pygetwindow as gw
from prettytable import PrettyTable
from datetime import datetime
import config
import time
import signal
import sys
import json
import xlsxwriter

class event: 
	def __init__(self, title, date): 
		self.title = title 
		self.first = date
		self.last = date
		self.count = 1

events = []
windows = []

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
	print(f">>> Elapsed time: {GetTimer(end-start)}\n")

def GetTimer(sec):
	hours, rem = divmod(sec, 3600)
	minutes, seconds = divmod(rem, 60)
	return "{:0>2}:{:0>2}:{:05.4f}".format(int(hours),int(minutes),seconds)

def getDuration(value):
	count = 0.90 / config.interval
	seconds = value / count
	return GetTimer(seconds)

def status():
	end = time.time()
	full = 0
	for e in events:
		full = full + e.count
	events_table = PrettyTable(["FIRST","LAST","DURATION","%","TITLE"])
	sorted_events = sorted(events, key=lambda x: x.count, reverse=True)

	# JSON Dump
	t = datetime.now().strftime("%Y%m%d")
	with open(f"data/dump_{t}.json", "w") as outfile:
		json.dump([obj.__dict__ for obj in sorted_events], outfile)

	# Excel report
	workbook = xlsxwriter.Workbook(f"data/report_{t}.xlsx")
	worksheet = workbook.add_worksheet("Data")
	bold = workbook.add_format({"bold": True})
	number_format = workbook.add_format({"num_format": "##0.00"})
	worksheet.write("A1", "FIRST", bold)
	worksheet.write("B1", "LAST", bold)
	worksheet.write("C1", "%", bold)
	worksheet.write("D1", "TITLE", bold)
	row = 1
	col = 0
	total_p = 0.0

	for e in sorted_events:
		p = percentage(full, e.count)
		w = e.title[-config.title_size:]
		if p > config.ignore_event:
			events_table.add_row([e.first, e.last, getDuration(e.count), p, w])
			worksheet.write(row, col, e.first)
			worksheet.write(row, col+1, e.last)
			worksheet.write_number(row, col+2, p, number_format)
			worksheet.write_string(row, col+3, w)
			row += 1
			total_p += p

	if total_p < 100.0:
		others = 100.0 - total_p
		worksheet.write(row, col, "N/A")
		worksheet.write(row, col+1, "N/A")
		worksheet.write_number(row, col+2, others, number_format)
		worksheet.write_string(row, col+3, "Others")
		row += 1

	sys.stdout.write("\r" + str(events_table) + "\n")
	sys.stdout.flush()
	elapsedTime(start, end)

	pie_chart = workbook.add_chart({"type":"pie"})
	pie_chart.set_legend({"none": True})
	pie_chart.add_series({
		"name":"My productivity today",
		"categories":f"=Data!$D$2:$D${row}",
		"values":f"=Data!$C$2:$C${row}",
		"data_labels":{"value":True,"category_name":True,"position":"outside_end"}
    })
	pie_chart.set_style(10)
	pie_chart.set_legend({"position": "right"})
	pie_chart.set_size({"width": 720, "height": 560})
	worksheet.insert_chart("H3", pie_chart)
	workbook.close()

def signal_handler(signal, frame):
	print("\033c")
	print(f"SUMMARY\n")
	status()
	sys.exit(0)

def animation(value):
	animation = "|/-\\"
	sys.stdout.write("\r" + animation[value % len(animation)])
	sys.stdout.flush()

if __name__ == "__main__":
	signal.signal(signal.SIGINT, signal_handler)
	print("[INFO] Start monitoring, press CTRL + C to stop")
	progress = 0
	time.sleep(1)
	start = time.time()

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