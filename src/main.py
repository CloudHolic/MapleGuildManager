import os
from datetime import date
from openpyxl import Workbook
from Crawler import Crawler


def printProgressBar (iteration, total, prefix='', suffix='',
                      decimals=1, length=100, fill='█', print_end='\r'):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filled_length = int(length * iteration // total)
    bar = fill * filled_length + '-' * (length - filled_length)
    print(f'\r{prefix} |{bar}| {percent}% {suffix}', end=print_end)
    if iteration == total:
        print()


if __name__ == "__main__":
    server_name = "scania"
    guild_name = "아이엠캔들"
    today = date.today().strftime('%m%d')

    members = Crawler.get_members(server_name, guild_name)
    write_wb = Workbook()
    write_ws = write_wb.active

    idx = 0
    printProgressBar(idx, len(members), prefix='Progress:', suffix='Complete', length=50)
    for member in members:
        write_ws.append([member[0], member[1], member[2], Crawler.get_mulung(member[0]), member[3]])
        idx += 1
        printProgressBar(idx, len(members), prefix='Progress:', suffix='Complete', length=50)
    write_wb.save(os.getcwd() + "/" + guild_name + "길드원" + today + "list.xlsx")
