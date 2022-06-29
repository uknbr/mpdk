"""
Monitor productivity
Home Office
"""

from sys import stdout
from datetime import datetime
import wx.adv
import wx
import logging
import threading
import time
import os
import mpdk


""" System tray """
tray = {"tooltip": "MPDK - My Productivity Dont Keel", "icon": "icon.png"}

""" Menu icons """
menu_icon = {
    "status": "menu1.png",
    "about": "menu2.png",
    "exit": "menu3.png",
}

def get_logger(
    name: str,
    log_file: str = None,
    level=logging.DEBUG,
    formatter=logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"),
):
    """
    Handlers for file and stdout
    Debug level as default
    """
    consoleHandler = logging.StreamHandler(stdout)
    consoleHandler.setLevel(level)
    consoleHandler.setFormatter(formatter)

    fileHandler = logging.FileHandler("mpdk.log")
    fileHandler.setLevel(level)
    fileHandler.setFormatter(formatter)

    logger = logging.getLogger(name)
    logger.setLevel(level)

    logger.addHandler(consoleHandler)
    logger.addHandler(fileHandler)

    return logger


logger = get_logger(__file__)


""" Available status of event_history """
event_name = {"init": 0, "start": 1, "lunch_1": 2, "lunch_2": 3, "end": 4}


def create_menu_item(menu, label, ico, func):
    """ Create item on menu with icon """

    item = wx.MenuItem(menu, -1, label)
    item.SetBitmap(wx.Bitmap(ico))
    menu.Bind(wx.EVT_MENU, func, id=item.GetId())
    menu.Append(item)
    return item


def show_popup(msg, flag):
    """ Windows notification """
    nmsg = wx.adv.NotificationMessage(title=tray["tooltip"], message=msg)
    nmsg.SetFlags(flag)
    nmsg.Show(timeout=wx.adv.NotificationMessage.Timeout_Auto)


def show_message(msg):
    """ Classic Message box """
    wx.MessageBox(msg, tray["tooltip"], wx.OK | wx.ICON_INFORMATION)


class CountingThread(threading.Thread):
    """ Thread to count time and trigger alert """

    def __init__(self, parent, period, msg):
        threading.Thread.__init__(self)
        self._parent = parent
        self._period = period
        self._msg = msg

    def run(self):
        logger.debug(f"Wait for {self._period} until next pop-up")
        #time.sleep(self._period)
        logger.debug(self._msg)
        #show_popup(self._msg, wx.ICON_WARNING)
        mpdk.init()


class TaskBarIcon(wx.adv.TaskBarIcon):
    def __init__(self, frame):
        self.frame = frame
        super(TaskBarIcon, self).__init__()
        self.set_icon(tray["icon"])
        self.Bind(wx.adv.EVT_TASKBAR_LEFT_DOWN, self.on_start)

    def CreatePopupMenu(self):
        menu = wx.Menu()
        create_menu_item(menu, "Status", menu_icon["status"], self.on_status)
        menu.AppendSeparator()
        create_menu_item(menu, "About", menu_icon["about"], self.on_about)
        create_menu_item(menu, "Exit", menu_icon["exit"], self.on_exit)
        return menu

    def set_icon(self, path):
        icon = wx.Icon(path)
        self.SetIcon(icon, tray["tooltip"])

    def on_status(self, event):
        logger.debug("Status")
        msg = "status"
        show_message(msg)

    def on_start(self, event):
        logger.info("Hello!")
        msg = "test"
        show_message(msg)
        wait_for = 60
        worker = CountingThread(self, period=wait_for, msg="See you tomorrow!")
        worker.start()

    def on_about(self, event):
        msg = "test"
        show_message(msg)

    def on_exit(self, event):
        wx.CallAfter(self.Destroy)
        self.frame.Close()
        logger.info("Bye!")
        os._exit(0)


class App(wx.App):
    def OnInit(self):
        frame = wx.Frame(None)
        self.SetTopWindow(frame)
        TaskBarIcon(frame)
        return True


def main():
    logger.info("Starting...")
    app = App(False)
    app.MainLoop()


if __name__ == "__main__":
    main()
