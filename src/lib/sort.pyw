# -*- coding: utf-8 -*-
import sys
import os
from datetime import datetime as dt

def sort_mru():
    #TODO 引数でvbからもらう事にする｡
    filename = os.environ["HOME"] + "\\Dropbox\\functional\\synced_setting_files\\Office\\Excel\\VBA2\\.cache\\mru.txt"

    buf = []
    with open(filename, "r") as f1:
        lines = f1.readlines()
        lines_sorted = sorted(lines, key=lambda line: line.decode("cp932").split(":::")[1], reverse=True)
        lines_sorted = sorted(lines, key=lambda line: get_open_date(line), reverse=True)
        buf = lines_sorted
    with open(filename, "w") as f2:
        for l in buf:
            l = l.decode("cp932")
            f2.write(l.encode("cp932"))

def get_open_date(s):# {{{
    s = s.decode("cp932").split(":::")[2].replace("\n","")
    try:
        s = dt.strptime(s, '%Y/%m/%d %H:%M:%S')
    except:
        s = dt.strptime('2015/04/01 9:00:00','%Y/%m/%d %H:%M:%S')
    return s# }}}

# if __name__ == __main__: command promptから呼び出されるため。
sort_mru()
