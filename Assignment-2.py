import psycopg2
from openpyxl.workbook import Workbook
import pandas as pand

class Total_compensation:
    def compensation(self):
        try:
            # to retrieve table from postgresql
            conn = psycopg2.connect(
                host="localhost",
                database="my_assignment",
                user="postgres",
                password="Krantideep@1")
            # Creating a cursor object using the cursor() method
            cur = conn.cursor()
            # we are reading our table which  we imported using connection through querry
            script_querry = """select emp.ename, emp.empno, dept.dname, (case when enddate is not null then ((enddate-startdate+1)/30)*(jobhist.sal) else ((current_date-startdate+1)/30)*(jobhist.sal) end)as Total_Compensation,
(case when enddate is not null then ((enddate-startdate+1)/30) else ((current_date-startdate+1)/30) end)as Months_Spent from jobhist, dept, emp 
where jobhist.deptno=dept.deptno and jobhist.empno=emp.empno"""
            cur.execute(script_querry)
            #cur.execute('select empno from jobhist')



            column_name = [desc[0] for desc in cur.description]
            data = cur.fetchall()
            new_file = pand.DataFrame(list(data), columns=column_name)

            file = pand.ExcelWriter('Assignment-2.py.xlsx')
            new_file.to_excel(file, sheet_name='bar')
            file.save()
# throw exception
        except Exception as e:
            print("Something went wrong", e)
        finally:

            if conn is not None:
                cur.close()
                conn.close()


if __name__=='__main__':
    conn = None
    cur = None
    comp = Total_compensation()
    comp.compensation()










