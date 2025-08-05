import sqlite3
from sqlite3 import Error
import random
from prettytable import PrettyTable


def create_table(db_file):
    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()

    columns = []
    for letter in 'ABCDEFGH':  # A-H列
        for num in range(1, 13):  # 1-12行
            columns.append(f"{letter}{num} REAL")

    # 创建数据表
    sql_create_table = f"""
    CREATE TABLE IF NOT EXISTS imported_data (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        original_file_name TEXT,
        updated_time TIMESTAMP DEFAULT (datetime('now', 'localtime')),
        {', '.join(columns)}
    )
    """

    try:
        # 执行创建表
        cursor.execute(sql_create_table)
        conn.commit()
        print("表创建成功，包含A1-H12共96列")
        
        # 验证表结构
        cursor.execute("PRAGMA table_info(imported_data)")
        columns_info = cursor.fetchall()
        print(f"表包含 {len(columns_info)-1} 个数据列(不含ID主键)")
        
    except sqlite3.Error as e:
        print(f"创建表时出错: {e}")
    finally:
        # 关闭连接
        cursor.close()
        conn.close()


def insert_from_list(data_list, db_file, original_file_name):
    """
    从列表数据源插入数据
    :param data_list: 二维列表，每个子列表包含96个元素对应A1-H12
    :param db_name: 数据库文件名
    """
    #data_list = data_list[0]
    append_data = [original_file_name] + data_list
    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()
    
    # 生成列名列表
    columns = []
    for letter in 'ABCDEFGH':
        for num in range(1, 13):
            columns.append(f"{letter}{num}")
    columns = ["original_file_name"] + columns
    
    # 准备SQL语句
    placeholders = ', '.join(['?'] * len(columns))

    sql = f"INSERT INTO imported_data ({', '.join(columns)}) VALUES ({placeholders})"
    # 批量插入
    cursor.executemany(sql, [append_data])
    #cursor.execute("INSERT INTO imported_data (original_file_name) VALUES (?)", original_file_name)
    conn.commit()
    print(f"成功插入 {len(data_list)} 条记录")
    conn.close()

def get_sql_list(db_file, table_name, columns):
    """
    读取SQLite数据库表中指定列的数据
    :param db_file: 数据库文件路径
    :param table_name: 表名
    :param columns: 要查询的列名列表
    :return: 查询结果列表
    """
    conn = None
    try:
        # 连接数据库
        conn = sqlite3.connect(db_file)
        cursor = conn.cursor()
        
        # 构建查询语句
        query = f"SELECT {', '.join(columns)} FROM {table_name}"
        
        # 执行查询
        cursor.execute(query)
        
        # 获取所有结果
        results = cursor.fetchall()
        
        # 获取列名
        column_names = [desc[0] for desc in cursor.description]
        
        #print(column_names, results)
        
        # 创建PrettyTable对象
        table = PrettyTable()
        table.field_names = columns
        table.align = "l"  # 左对齐
        
        # 添加数据到表格
        for row in results:
            table.add_row(row)

        print(table, len(results))
        target_id = input("choose a id:")

        # 构建查询语句
        query = f"SELECT {', '.join(columns)} FROM {table_name} WHERE id = ?"
        
        # 执行查询
        cursor.execute(query, (target_id,))
        
        # 获取所有结果
        results = cursor.fetchall()

        # 创建PrettyTable对象
        table = PrettyTable()
        table.field_names = columns
        table.align = "l"  # 左对齐
        
        # 添加数据到表格
        for row in results:
            table.add_row(list(row))
        print(table)

        # 生成列名列表
        columns = []
        for letter in 'ABCDEFGH':
            for num in range(1, 13):
                columns.append(f"{letter}{num}")

        query = f"SELECT {', '.join(columns)} FROM {table_name} WHERE id = ?"
        # 执行查询
        cursor.execute(query, (target_id,))
        
        # 获取所有结果
        results = cursor.fetchall()
        #print(list(results[0]))


        return results[0]
        
    except sqlite3.Error as e:
        print(f"SQLite错误: {e}")
        return None, None
    finally:
        if conn:
            conn.close()

def generate_sample_data(rows):
    # 生成示例数据：每行96个0-100的随机浮点数
    out = [[random.uniform(0, 100) for _ in range(96)] for _ in range(rows)]
    return out[0]

if __name__ == '__main__':
    path = r"D:\test.db"

    get_sql_list(path, "imported_data", ["id", "original_file_name", "updated_time"])

    #create_table(path)

    #insert_from_list(generate_sample_data(1), path, original_file_name = "20250726.xls")