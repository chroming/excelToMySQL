# -*- coding:utf-8 -*-

import xlrd
import MySQLdb


def export_to_sql(filename):
    year = int(filename[0:4])
    db = MySQLdb.connect("localhost", "root", "", "zhaosheng", charset="utf8")
    cursor = db.cursor()
    xlsfile = xlrd.open_workbook(filename)
    xlssheet = xlsfile.sheets()[0]
    nrows = xlssheet.nrows
    ncols = xlssheet.ncols
    for r in range(2, nrows):
        for c in range(2, ncols, 3):
            local = xlssheet.cell(0, c).value
            spe_name = xlssheet.cell(r, 0).value
            course = xlssheet.cell(r, 1).value
            low = xlssheet.cell(r, c+1).value
            high = xlssheet.cell(r, c).value
            aver = xlssheet.cell(r, c+2).value
            insert_sql = "insert into fenshuxian values ('%s', '%s', '%s', '%s', '%s', '%s', '%s')" % (year, local, spe_name, course, low, high, aver)
            print insert_sql
            cursor.execute(insert_sql)
            db.commit()
            #raw_input()
            #print xlssheet.cell(r, c).value
    db.close()

filelist = []

for filename in filelist:
    export_to_sql(filename)
