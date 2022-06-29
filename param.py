from datetime import datetime


def data_init():
    return []


def time_start():
    return 0


date_today = datetime.now().strftime("%Y%m%d")
folder_data = "data"
file_json = f"{folder_data}/dump_{date_today}.json"
file_excel = f"{folder_data}/report_{date_today}.xlsx"
