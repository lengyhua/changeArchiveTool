import vertica_python

import info
import xsltool

# 档案表
ARCHIVE_TABLE = 'viid_facestatic.people_archive'
# 轨迹表
TRACK_TABLE = 'viid_facestatic.people_track'
# 人脸抓拍表
SNAP_FACE_TABLE_5 = 'viid_facesnap.facesnapstructured_a050000'
SNAP_FACE_TABLE_0 = 'viid_facesnap.facesnapstructured_a000000'


# 连接vertica
def connect_vertica(xls_file='info.xlsx'):
    workbook = xsltool.read_xlsx(xls_file)
    vertica_connect_info = xsltool.read_db_info(workbook)
    return vertica_python.connect(host=vertica_connect_info[0],
                                  port=vertica_connect_info[1],
                                  user=vertica_connect_info[2],
                                  password=vertica_connect_info[3],
                                  database=vertica_connect_info[4])


# 指定档案ID删除档案
def remove_archive(vertica_conn, delete_info):
    if not delete_info or not vertica_conn:
        return
    archive_id = delete_info.value
    if not archive_id or not archive_id.strip():
        return
    archive_id = archive_id.strip()
    cur = vertica_conn.cursor()
    print("delete archive: %s" % archive_id)
    try:
        cur.execute("delete from %s where record_id = '%s'" % (ARCHIVE_TABLE, archive_id))
        cur.execute("delete from %s where people_id = '%s'" % (TRACK_TABLE, archive_id))
        vertica_conn.commit()
    except:
        print("error delete archive: %s, rollback now" % archive_id)
        delete_info.result = info.FAIL
        vertica_conn.rollback()


# 修改档案封面
def update_archive_info(vertica_conn, update_info):
    if not vertica_conn or not update_info:
        return
    values = update_info.values
    cur = vertica_conn.cursor()
    archive_id, repo_id, id_number, id_type, name, gender, phone, url, face_id = values
    if not repo_id and not id_number and not id_type and not name and not gender and not phone and not url and not face_id:
        return
    print("change archive info: %s" % archive_id)
    try:
        sql = "update %s set " % ARCHIVE_TABLE
        set_sql = []
        if repo_id:
            set_sql.append("repo_id = '%s'" % repo_id)
        if id_number:
            set_sql.append("id_number = '%s'" % id_number)
        if id_type:
            set_sql.append("id_type = %s" % id_type)
        if name:
            set_sql.append("name = '%s'" % name)
        if gender:
            set_sql.append("gender = %s" % gender)
        if phone:
            set_sql.append("phone = %s" % phone)
        if url:
            set_sql.append("cover_image = '%s'" % url)
        elif face_id:
            cur.execute("select imageurlpart from %s where faceid = '%s' " \
                        "limit 1" % (SNAP_FACE_TABLE_5, face_id))
            result = cur.fetchone()
            if not result:
                cur.execute("select imageurlpart from %s where faceid = '%s' " \
                            "limit 1" % (SNAP_FACE_TABLE_0, face_id))
                result = cur.fetchone()
            if result:
                set_sql.append("cover_image = '%s'" % result[0])
        if len(set_sql) == 0:
            update_info.result = info.ERROR_INFO
            return
        sql = sql + ",".join(set_sql) + "where record_id = '%s'" % archive_id
        cur.execute(sql)
        vertica_conn.commit()
    except:
        print("error change archive info: %s, rollback now" % archive_id)
        update_info.result = info.FAIL
        vertica_conn.rollback()


# 删除档案指定轨迹
def delete_track(vertica_conn, delete_info):
    if not vertica_conn or not delete_info:
        return
    track_id = delete_info.value
    if not track_id or not track_id.strip():
        return
    track_id = track_id.strip()
    print("delete track: %s" % track_id)
    cur = vertica_conn.cursor()
    try:
        cur.execute("delete from %s where snap_id = '%s'" % (TRACK_TABLE, track_id))
        vertica_conn.commit()
    except:
        print("error delete track: %s, rollback now" % track_id)
        delete_info.result = info.FAIL
        vertica_conn.rollback()


# 合并档案
def merge_archive(vertica_conn, merge_info):
    if not vertica_conn or not merge_info:
        return
    from_archive = merge_info.from_archive
    to_archive = merge_info.to_archive
    if not from_archive or not from_archive.strip() \
            or not to_archive or not to_archive.strip():
        merge_info.result = info.ERROR_INFO
        return
    if from_archive == to_archive:
        merge_info.result = info.ERROR_INFO
        return
    print("merge archive from %s to %s" % (from_archive, to_archive))
    cur = vertica_conn.cursor()
    try:
        cur.execute("select * from %s where record_id = '%s'" % (ARCHIVE_TABLE, to_archive))
        result = cur.fetchone()
        if not result:
            merge_info.result = info.NOT_FOUND
            return
        cur.execute("update %s set people_id = '%s' where people_id = '%s'" % (TRACK_TABLE,
                                                                               to_archive, from_archive))
        cur.execute("delete from %s where record_id = '%s'" % (ARCHIVE_TABLE, from_archive))
        vertica_conn.commit()
    except:
        merge_info.result = info.FAIL
        vertica_conn.rollback()
