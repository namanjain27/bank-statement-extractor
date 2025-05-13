
from bank_statement_parser import BankStatementParser

# def import_bank_statement("feb25.xlsx"):
parser = BankStatementParser()
transactions = parser.parse_statement(r"D:\codes\bank statement extractor\feb25.xlsx")
print(transactions)

# Then use transactions in your app