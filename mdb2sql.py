import pyodbc
import pymysql
from typing import NoReturn, Tuple, List


def conn_access(filepath: str) -> Tuple:
    """
    连接access文件，并返回连接和游标
    :param filepath: access文件绝对路径
    :return: (connect, cursor)
    """
    conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            r'DBQ=%s;' % filepath
    )
    cnxn = pyodbc.connect(conn_str)
    return cnxn, cnxn.cursor()


def get_info(cursor) -> List[Tuple[str, str]]:
    """
    获取access表名和表中的字段
    :param cursor： access游标
    :return: 表名和字段字符串
    """
    tables = [table_info.table_name for table_info in cursor.tables(tableType='TABLE')]
    tables_info = []
    for table in tables:
        table_columns = ''  # 字段字符串
        for row in cursor.columns(table):
            table_columns += '{} {}({}),'.format(row.column_name, row.type_name, row.column_size)
        table_columns = table_columns.rstrip(',')  # 去掉字符串末尾的,
        tables_info.append((table, table_columns))
    return tables_info


def conn_mysql(host: str, user: str, password: str, db: str) -> Tuple:
    """
    连接mysql数据库，返回连接和游标
    :param host: 主机名
    :param user: 用户
    :param password: 密码
    :param db: 数据库名
    :return: (connect, cursor)
    """
    connect = pymysql.connect(host, user, password, db)
    cursor = connect.cursor()
    return connect, cursor


def create_table(cursor, tables_info) -> NoReturn:
    """
    在mysql中创建对应的数据表
    :param cursor: 游标
    :param tables_info: 从access中获取的表信息
    :return: None
    """
    for table in tables_info:
        sql = 'CREATE TABLE IF NOT EXISTS {}({})'.format(table[0], table[1])
        cursor.execute(sql)


def add_rows(tables_info, cursor, cursor1, connect1) -> NoReturn:
    """
    插入数据到数据表中
    :param connect1: mysql连接
    :param tables_info: 从access中获取的表信息
    :param cursor: access游标
    :param cursor1: mysql游标
    :return: None
    """
    for table in tables_info:
        cursor.execute('SELECT * FROM {}'.format(table[0]))
        rows = cursor.fetchall()
        rows_list = []
        for row in rows:
            rows_list.append(tuple(row))

        length_row = len(rows_list[0])
        str1 = ('%s, ' * length_row).rstrip(', ')
        insert_sql = 'INSERT INTO {} VALUES ({})'.format(table[0], str1)
        cursor1.executemany(insert_sql, rows_list)
        connect1.commit()


def close(connect, cursor, connect1, cursor1) -> NoReturn:
    """
    关闭数据库连接
    :param connect: access连接
    :param cursor: access游标
    :param connect1: mysql连接
    :param cursor1: mysql游标
    :return:
    """
    cursor.close()
    connect.close()
    cursor1.close()
    connect1.close()


def mdb2sql(filepath, host, user, password, db) -> NoReturn:
    """
    将 .mdb 文件数据插入mysql中
    :param filepath: .mdb文件绝对路径
    :param host: mysql主机名
    :param user: mysql用户名
    :param password: mysql密码
    :param db: mysql数据库
    :return: None
    """
    connect, crsr_access = conn_access(filepath)
    tables_info = get_info(crsr_access)
    connect1, crsr_mysql = conn_mysql(host, user, password, db)
    create_table(crsr_mysql, tables_info)
    add_rows(tables_info, crsr_access, crsr_mysql, connect1)
    close(connect, crsr_access, connect1, crsr_mysql)
    print('Done！！！')
