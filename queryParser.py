from connectionExecution import connectionResults

headerFooterResult, rowResult = connectionResults(10)
reshapedHeaderFooterData = []
for item in headerFooterResult:
    _, description, company, code, _, start_date, end_date, primary_id, * \
        decimals, status, _ = item
    decimals = [float(d) for d in decimals]

    reshapedHeaderFooterData.append({
        'Description': description,
        'Company': company,
        'Code': code,
        'StartDate': start_date,
        'EndDate': end_date,
        'PrimaryID': primary_id,
        'Decimals': decimals,
        'Status': status
    })

# {
#   'Description': 'فاتورة ضريبية مبسطة',
#   'Company': 'مؤسسة تسويق البناء - شقراء',
#   'Code': 'D0040',
#   'StartDate': datetime.datetime(2019, 5, 5, 0, 0),
#   'EndDate': datetime.datetime(2019, 5, 20, 0, 0),
#   'PrimaryID': 'Primary100009',
#   'Decimals': [773.4, 25.0, 193.35, 580.05, 29.0, 609.05],
#   'Status': 'S10'
# }

# Print the reshaped data
for item in reshapedHeaderFooterData:
    print(item)
