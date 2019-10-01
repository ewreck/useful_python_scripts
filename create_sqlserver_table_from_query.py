import pymssql

conn = pymssql.connect(server= srvr, user = usr, password=pwd, database=db)
cursor = conn.cursor()

cursor.execute(
'''
SELECT *
FROM table;
''')
results  = cursor.fetchall()

def translate_datatype(x):
    if x == 1: return 'VARCHAR(255)'
    if x == 2: return 'BIT'
    if x == 3: return 'INT'
    if x == 4: return 'DATETIME'
    if x == 5: return 'ROWID'
    else: raise ValueError
def create_var_string(x):
    if x == 1 or x == 4: return '%s'
    if x == 2 or x == 3 or x == 5: return '%d'
    else: raise ValueError

create_table_str = ""
col_names = ""
value_input = ""
for x in range(len(cursor.description)):
    desc = cursor.description[x]
    if x == 0:
        create_table_str = str(desc[0]) + ' ' + translate_datatype(desc[1])
        col_names = str(desc[0])
        value_input = create_var_string(desc[1])
    else:
        create_table_str += ', ' + str(desc[0]) + ' ' + translate_datatype(desc[1])
        col_names += ', ' + str(desc[0])
        value_input += ', ' + create_var_string(desc[1])

print(create_table_str)
print(col_names)
print(value_input)

cursor.execute(
'''
IF OBJECT_ID('test_table_name') IS NOT NULL DROP TABLE test_table_name;
CREATE TABLE test_table_name (%s);
''' % (create_table_str)
)
conn.commit()
