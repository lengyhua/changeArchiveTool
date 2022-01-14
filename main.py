import os.path
import sys

import dbtools
import xsltool
from dbtools import connect_vertica


# 删除档案
def remove_archive(xls='info.xlsx'):
    conn = connect_vertica(xls)
    workbook = xsltool.read_xlsx(xls)
    delete_archives = xsltool.read_delete_archive(workbook)
    for delete_archive in delete_archives:
        dbtools.remove_archive(conn, delete_archive)
        xsltool.write_delete_archive_result(workbook, delete_archive, xls)


# 删除轨迹
def remove_tracks(xls='info.xlsx'):
    conn = connect_vertica(xls)
    workbook = xsltool.read_xlsx(xls)
    delete_tracks = xsltool.read_delete_tracks(workbook)
    for delete_track in delete_tracks:
        dbtools.delete_track(conn, delete_track)
        xsltool.write_delete_track_result(workbook, delete_track, xls)


# 合并档案信息
def merge_archive(xls='info.xlsx'):
    conn = connect_vertica(xls)
    workbook = xsltool.read_xlsx(xls)
    merge_archives = xsltool.read_merge_archive(workbook)
    for m in merge_archives:
        dbtools.merge_archive(conn, m)
        xsltool.write_merge_archive_result(workbook, m, xls)


# 更新档案信息
def update_archive(xls='info.xlsx'):
    conn = connect_vertica()
    workbook = xsltool.read_xlsx(xls)
    update_archives = xsltool.read_update_archive(workbook)
    for archive in update_archives:
        dbtools.update_archive_info(conn, archive)
        xsltool.write_update_archive_result(workbook, archive, xls)


if __name__ == '__main__':
    xls_file = 'info.xlsx'
    if len(sys.argv) > 1:
        if os.path.exists(sys.argv[1]) and os.path.isfile(sys.argv[1]):
            xls_file = sys.argv[1]
    remove_archive(xls_file)
    remove_tracks(xls_file)
    update_archive(xls_file)
    merge_archive(xls_file)
