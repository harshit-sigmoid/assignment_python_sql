import psycopg2
from openpyxl.workbook import Workbook
import pandas as pand


class employees_data:
    def emp(self):
        try:
            # to retrieve table from postgresql
            connection = psycopg2.connect(
                host="localhost",
                database="my_assignment",
                user="postgres",
                '''
                sharing of confidential data should be avoided
                use separate file for storing them 
                or use environment variables
                such as env.password
                to store such informations
                '''
                password="Krantideep@1") 
            # Creating a cursor object using the cursor() method
            cur = connection.cursor()


            # we are reading our table which  we imported using connection through querry
            '''could be named in format such as query_<type>'''
            script = """SELECT e1.empno, e1.ename, (case when mgr is not null then (select ename from emp as e2 where e1.mgr=e2.empno limit 1) else null end) as manager
                        from emp as e1"""
            #cur.execute(' to select empno from jobhist table')
            cur.execute(script)


            column_name = [desc[0] for desc in cur.description]
            file_name = cur.fetchall()
            '''file is still a dataframe, could be named as final_table'''
            file = pand.DataFrame(list(file_name), columns=column_name)

            Creating_xlsx = pand.ExcelWriter('ansno_1.py.xlsx')
            file.to_excel(Creating_xlsx, sheet_name='bar')
            Creating_xlsx.save()

        except Exception as exc:
            print(" Sorry Error Occured",exc)

        finally:

            if connection is not None:
                cur.close()
                connection.close()


if __name__ == '__main__':
    connection = None
    cur = None
    # creating object of employee_data class
    employee = employees_data()
    employee.emp()
