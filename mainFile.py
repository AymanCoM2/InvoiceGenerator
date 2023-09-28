from connectionExecution import fillResults

headerFooterResult , rowResult = fillResults(50)

for hfRow in headerFooterResult:
    print(hfRow) # Prints the Tuple Of the Data 

print("----------------------------") # Seperation Of Rows From Headers and Footer
for rRow in rowResult:
    print(rRow) ## Print Each Tuple Of Items In FATOORA 
    print("****************")
