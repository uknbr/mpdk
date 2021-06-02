#!/usr/bin/env python3
import pygetwindow as gw
from prettytable import PrettyTable
from datetime import datetime
import config
import time
import signal
import sys

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

def timer(start, end):
	hours, rem = divmod(end-start, 3600)
	minutes, seconds = divmod(rem, 60)
	print(">>> Elapsed time {:0>2}:{:0>2}:{:05.4f}\n".format(int(hours),int(minutes),seconds))

def status():
	end = time.time()
	full = 0
	for e in events:
		full = full + e.count
	events_table = PrettyTable(["FIRST","LAST","COUNT","%","TITLE"])
	sorted_events = sorted(events, key=lambda x: x.count, reverse=True)
	for e in sorted_events:
		p = percentage(full, e.count)
		if p > config.ignore_event:
			events_table.add_row([e.first, e.last, e.count, p, e.title[-config.title_size:]])
	sys.stdout.write("\r" + str(events_table) + "\n")
	sys.stdout.flush()
	timer(start, end)

def signal_handler(signal, frame):
	print("\033c")
	print(f"===== SUMMARY =====\n")
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
			dt = datetime.now().strftime("%c")
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
