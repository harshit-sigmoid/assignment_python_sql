import psycopg2
from openpyxl.workbook import Workbook
import pandas as pand
# creating employee class

''' follow camel casing wile naming Classes   eg.  class employeesData'''
class employees_data:

    def emp(self):
        # try method
        try:
            conn = psycopg2.connect(
                # to retrieve table from postgresql
                database="my_assignment",
                user="postgres",
                password="Krantideep@1")
            # Creating a cursor object using the cursor() method
            cursor = conn.cursor()
            # we are reading our table which  we imported using connection through querry
            querry = """
                    select dept.deptno, dept_name, sum(total_compensation) from Compensation, dept
                    where Compensation.dept_name=dept.dname
                    group by dept_name, dept.deptno
                    """
             # executing querry
            cursor.execute(querry)

            column_name = [desc[0] for desc in cursor.description]
            file = cursor.fetchall()
            new_file = pand.DataFrame(list(file), columns=column_name)
             # converting .py to xlsx
            changed_file = pand.ExcelWriter('Assignment_4.py_4.xlsx')
            new_file.to_excel(changed_file, sheet_name='bar')
            changed_file.save()


        except Exception as exc:
            print("try again", exc)

        finally:

            if conn is not None:
                cursor.close()
                conn.close()


if __name__ == '__main__':
    conn = None
    cur = None
    # creating object
    employee = employees_data()
    employee.emp()
