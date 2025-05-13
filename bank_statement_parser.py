import pandas as pd
import numpy as np
import re
from datetime import datetime

class BankStatementParser:
    """
    A class to parse bank statements from Excel files with automatic field detection.
    Works with statements from different banks by identifying the transaction table
    and key fields (date, description, withdrawal, deposit).
    """
    
    def __init__(self):
        # Common headers for each field type across different banks
        self.date_headers = [
            'date', 'transaction date', 'post date', 'trans date', 'value date', 
            'trade date', 'posting date', 'transaction time', 'date of transaction'
        ]
        
        self.description_headers = [
            'description', 'narration', 'details', 'particulars', 'transaction details',
            'remarks', 'payee', 'payment details', 'transaction description', 'narrative',
            'merchant', 'transaction', 'payment name', 'name', 'beneficiary', 'reference'
        ]
        
        self.withdrawal_headers = [
            'withdrawal', 'debit', 'debit amount', 'amount debit', 'payments', 
            'paid out', 'outgoing', 'money out', 'withdrawals', 'dr', 'dr amount',
            'spend', 'expense', 'payment', 'withdraw amount', 'withdrawal amount',
            'amount withdrawn', 'debit(dr)', 'debits'
        ]
        
        self.deposit_headers = [
            'deposit', 'credit', 'credit amount', 'amount credit', 'deposits', 
            'paid in', 'incoming', 'money in', 'cr', 'cr amount', 'income',
            'deposit amount', 'credit(cr)', 'amount deposited', 'credits',
            'inflow', 'money in', 'received'
        ]
        
        # Regular expressions for detecting dates
        self.date_patterns = [
            r'\d{2}[/\-\.]\d{2}[/\-\.]\d{2,4}',  # DD/MM/YYYY, DD-MM-YYYY, etc.
            r'\d{4}[/\-\.]\d{2}[/\-\.]\d{2}',    # YYYY/MM/DD, YYYY-MM-DD, etc.
            r'\d{1,2}\s+[A-Za-z]{3,}\s+\d{2,4}'  # DD Month YYYY
        ]

    def find_transaction_table(self, excel_file):
        """
        Locates the transaction table in an Excel file by examining all sheets
        and looking for date patterns and key headers.
        
        Returns:
            tuple: (DataFrame of the transaction table, sheet name, starting row)
        """
        try:
            # Read all sheets to examine them
            xls = pd.ExcelFile(excel_file)
            
            best_match = None
            best_score = 0
            
            # Check each sheet
            for sheet_name in xls.sheet_names:
                # Try reading with no header first
                df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
                
                # Look for date columns and key headers
                for start_row in range(min(20, len(df))):  # Check first 20 rows
                    # Create a view starting from this row
                    potential_table = df.iloc[start_row:].reset_index(drop=True)
                    
                    # Skip if empty
                    if len(potential_table) < 2:
                        continue
                    
                    # Calculate score based on headers and date patterns
                    score = self._calculate_table_score(potential_table)
                    
                    if score > best_score:
                        # Use the row as header and create proper DataFrame
                        header_df = pd.read_excel(
                            excel_file, 
                            sheet_name=sheet_name, 
                            header=start_row
                        )
                        best_match = (header_df, sheet_name, start_row)
                        best_score = score
            
            if best_match:
                return best_match
            
            # If no suitable table found, return the first sheet as fallback
            df = pd.read_excel(excel_file, sheet_name=0)
            return (df, xls.sheet_names[0], 0)
            
        except Exception as e:
            raise Exception(f"Error finding transaction table: {str(e)}")

    def _calculate_table_score(self, df):
        """Calculate a score indicating how likely this DataFrame contains transaction data."""
        score = 0
        
        # Check for header matches in the first row
        if not df.empty:
            first_row = df.iloc[0].astype(str).str.lower()
            
            # Check for date header matches
            for header in self.date_headers:
                if any(first_row.str.contains(header, regex=False)):
                    score += 10
                    break
            
            # Check for description header matches
            for header in self.description_headers:
                if any(first_row.str.contains(header, regex=False)):
                    score += 10
                    break
            
            # Check for withdrawal/deposit header matches
            withdrawal_found = False
            for header in self.withdrawal_headers:
                if any(first_row.str.contains(header, regex=False)):
                    score += 10
                    withdrawal_found = True
                    break
            
            deposit_found = False
            for header in self.deposit_headers:
                if any(first_row.str.contains(header, regex=False)):
                    score += 10
                    deposit_found = True
                    break
            
            # Extra points if both withdrawal and deposit columns are found
            if withdrawal_found and deposit_found:
                score += 15
        
        # Check for date patterns in the data
        for col in range(min(10, len(df.columns))):
            if df[col].dtype == object:  # Only check string columns
                # Convert to string and check first 10 rows for date patterns
                col_data = df[col].astype(str).head(10).str.lower()
                date_matches = sum(col_data.str.contains('|'.join(self.date_patterns), regex=True))
                if date_matches >= 3:  # If at least 3 rows match date patterns
                    score += 20
                    break
        
        # Check for numeric data (likely amounts)
        numeric_cols = 0
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]) or df[col].astype(str).str.match(r'^[-+]?\d*\.?\d+$').sum() > len(df) * 0.7:
                numeric_cols += 1
        
        # More numeric columns suggest financial data
        if numeric_cols >= 2:
            score += 10
        
        return score

    def identify_fields(self, df):
        """
        Identifies which columns correspond to date, description, withdrawal, and deposit.
        
        Returns:
            dict: Mapping of field types to column names
        """
        field_mapping = {
            'date': None,
            'description': None,
            'withdrawal': None,
            'deposit': None
        }
        
        # Normalize headers - convert all to lowercase
        df.columns = [str(col).lower().strip() for col in df.columns]
        
        # Replace NaN with empty string in headers
        headers = {str(col).lower(): col for col in df.columns}
        
        # First try direct header matching
        for header_name, header_variants in [
            ('date', self.date_headers),
            ('description', self.description_headers),
            ('withdrawal', self.withdrawal_headers),
            ('deposit', self.deposit_headers)
        ]:
            for variant in header_variants:
                matching_cols = [col for col in headers.keys() if variant in col]
                if matching_cols:
                    field_mapping[header_name] = headers[matching_cols[0]]
                    break
        
        # If date column not found by header, try to find by content pattern
        if field_mapping['date'] is None:
            for col in df.columns:
                if df[col].dtype == object:  # Check only string columns
                    date_pattern_matches = df[col].astype(str).str.match('|'.join(self.date_patterns)).sum()
                    if date_pattern_matches > len(df) * 0.7:  # If >70% of rows match date pattern
                        field_mapping['date'] = col
                        break
        
        # If description not found, choose the column with longest text values
        if field_mapping['description'] is None:
            text_cols = [col for col in df.columns if df[col].dtype == object]
            if text_cols:
                avg_lengths = {col: df[col].astype(str).str.len().mean() for col in text_cols}
                field_mapping['description'] = max(avg_lengths, key=avg_lengths.get)
        
        # For withdrawal/deposit, look for numeric columns if not found by header
        numeric_cols = [col for col in df.columns if pd.api.types.is_numeric_dtype(df[col]) or 
                        df[col].astype(str).str.replace(',', '').str.match(r'^[-+]?\d*\.?\d+$').sum() > len(df) * 0.5]
        
        # If only one numeric column, it might contain both withdrawals and deposits
        if len(numeric_cols) == 1 and (field_mapping['withdrawal'] is None or field_mapping['deposit'] is None):
            col = numeric_cols[0]
            # Try to convert to numeric, removing commas
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce')
            
            # Column with negative values = withdrawals, positive = deposits
            field_mapping['withdrawal'] = col
            field_mapping['deposit'] = col
        
        # If two or more numeric columns and fields still not found
        elif len(numeric_cols) >= 2:
            # Find columns that seem to have valid currency values
            currency_cols = []
            for col in numeric_cols:
                # Convert to numeric, removing commas
                df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce')
                if df[col].notna().sum() > len(df) * 0.5:  # More than 50% valid numbers
                    currency_cols.append(col)
            
            if len(currency_cols) >= 2 and (field_mapping['withdrawal'] is None or field_mapping['deposit'] is None):
                # If we have exactly two currency columns, assign them to withdrawal and deposit
                if len(currency_cols) == 2:
                    # The column with more negative values is likely withdrawals
                    neg_counts = [sum(df[col] < 0) for col in currency_cols]
                    
                    if field_mapping['withdrawal'] is None:
                        field_mapping['withdrawal'] = currency_cols[0] if neg_counts[0] >= neg_counts[1] else currency_cols[1]
                    
                    if field_mapping['deposit'] is None:
                        field_mapping['deposit'] = currency_cols[1] if neg_counts[0] >= neg_counts[1] else currency_cols[0]
                else:
                    # More than two currency columns - use header similarity as a tiebreaker
                    for col in currency_cols:
                        col_name = str(col).lower()
                        
                        # Check for withdrawal-like terms
                        if field_mapping['withdrawal'] is None and any(term in col_name for term in ['with', 'debit', 'out', 'dr']):
                            field_mapping['withdrawal'] = col
                        
                        # Check for deposit-like terms
                        if field_mapping['deposit'] is None and any(term in col_name for term in ['dep', 'credit', 'in', 'cr']):
                            field_mapping['deposit'] = col
                    
                    # If still not found, take the first two numeric columns
                    if field_mapping['withdrawal'] is None:
                        field_mapping['withdrawal'] = currency_cols[0]
                    
                    if field_mapping['deposit'] is None:
                        for col in currency_cols:
                            if col != field_mapping['withdrawal']:
                                field_mapping['deposit'] = col
                                break
        
        return field_mapping

    def parse_statement(self, excel_file):
        """
        Main method to parse a bank statement from Excel file.
        
        Returns:
            DataFrame: Standardized transaction data with date, description, withdrawal, deposit
        """
        # Find the transaction table
        df, sheet_name, start_row = self.find_transaction_table(excel_file)
        
        # Clean up the data - drop fully empty rows
        df = df.dropna(how='all')
        
        # Identify fields
        field_mapping = self.identify_fields(df)
        
        # Create standardized dataframe
        result_df = pd.DataFrame(columns=['Date', 'Description', 'Withdrawal', 'Deposit'])
        
        if field_mapping['date'] is not None:
            result_df['Date'] = df[field_mapping['date']]
            
            # Try to parse dates into a standard format
            try:
                result_df['Date'] = pd.to_datetime(result_df['Date'], errors='coerce')
                result_df['Date'] = result_df['Date'].dt.strftime('%Y-%m-%d')
            except:
                pass  # Keep as is if parsing fails
        
        if field_mapping['description'] is not None:
            result_df['Description'] = df[field_mapping['description']]
        
        # Handle withdrawal and deposit
        if field_mapping['withdrawal'] == field_mapping['deposit']:
            # Same column for both - negative values are withdrawals, positive are deposits
            col = field_mapping['withdrawal']
            try:
                values = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce')
                result_df['Withdrawal'] = values.where(values < 0, np.nan).abs()
                result_df['Deposit'] = values.where(values > 0, np.nan)
            except:
                result_df['Withdrawal'] = df[col]
                result_df['Deposit'] = df[col]
        else:
            if field_mapping['withdrawal'] is not None:
                try:
                    result_df['Withdrawal'] = pd.to_numeric(df[field_mapping['withdrawal']].astype(str).str.replace(',', ''), errors='coerce').abs()
                except:
                    result_df['Withdrawal'] = df[field_mapping['withdrawal']]
            
            if field_mapping['deposit'] is not None:
                try:
                    result_df['Deposit'] = pd.to_numeric(df[field_mapping['deposit']].astype(str).str.replace(',', ''), errors='coerce').abs()
                except:
                    result_df['Deposit'] = df[field_mapping['deposit']]
        
        # Clean up remaining NaN values
        result_df = result_df.fillna('')
        
        # Filter out rows without any transaction data
        result_df = result_df[~((result_df['Withdrawal'] == '') & (result_df['Deposit'] == ''))]
        
        # Filter out header-like rows that might have been included
        pattern = '|'.join(self.date_headers + self.description_headers + 
                          self.withdrawal_headers + self.deposit_headers)
        result_df = result_df[~result_df['Description'].astype(str).str.lower().str.contains(pattern, regex=True)]
        
        return result_df

# Example usage
def process_bank_statement(file_path):
    parser = BankStatementParser()
    try:
        transactions = parser.parse_statement(file_path)
        print(f"Successfully parsed {len(transactions)} transactions")
        return transactions
    except Exception as e:
        print(f"Error processing statement: {str(e)}")
        return None

# Example implementation for an expense tracker service
if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python bank_statement_parser.py <excel_file_path>")
        sys.exit(1)
    
    file_path = sys.argv[1]
    transactions = process_bank_statement(file_path)
    
    if transactions is not None:
        print("\nFirst 5 transactions:")
        print(transactions.head())
        
        # Here you could save to database, generate reports, etc.
        # Example: Save to CSV
        output_file = "parsed_transactions.csv"
        transactions.to_csv(output_file, index=False)
        print(f"\nTransactions saved to {output_file}")
