import pyodbc
from functions import replaceHeaderFooterQuery , replaceRowsQuery
def fillResults(docNumber):
    headerfooterQuery = replaceHeaderFooterQuery(str(docNumber))
    rowQuery  = replaceRowsQuery(str(docNumber))
    conn = pyodbc.connect("Driver={SQL Server};"
                    "Server=10.10.10.100;"
                    "Database=LB;"
                    "UID=ayman;"
                    "PWD=admin@1234;")
    cursor = conn.cursor()
    cursor.execute(headerfooterQuery)
    headerFooterResult  = cursor.fetchall()
    cursor.execute(rowQuery)
    rowResult  = cursor.fetchall()
    return headerFooterResult , rowResult