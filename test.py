
from statementExtractor import extract_transactions

transactions = extract_transactions(r"D:\codes\bank statement extractor\apr25.xls")
with open("output.txt", "a") as f:
    print(transactions, file=f)

# Then use transactions in your app