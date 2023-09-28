import pyodbc
from functions import replaceHeaderFooterQuery , replaceRowsQuery

conn = pyodbc.connect("Driver={SQL Server};"
                "Server=10.10.10.100;"
                "Database=LB;"
                "UID=ayman;"
                "PWD=admin@1234;")

headerfooterQuery = replaceHeaderFooterQuery(str(50))
rowQuery  = replaceRowsQuery(str(50))


cursor = conn.cursor()
cursor.execute(headerfooterQuery)
headerFooterResult  = cursor.fetchall()
cursor.execute(rowQuery)
rowResult  = cursor.fetchall()

print(headerFooterResult)
print("------------------------------------")
print(rowResult)


# Make Two Functions To Get the DocNum and Return the Query Concatenated With it 
# Make Functions To Run The Queries and Return the Results 
# Get Results and Parse Them 
# Loop and FIll the Tables and the PlaceHolders 

# 