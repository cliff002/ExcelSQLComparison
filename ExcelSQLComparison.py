class ExcelSQLComparison
    def __init__(self, excel_file, server, db, table)
        self.excel_file = excel_file
        self.server = server
        self.db = db
        self.table = table

    def compare(self, excel_columns, sql_columns)
        # Read Excel file and exclude first row
        df = pd.read_excel(self.excel_file, skiprows=1, usecols=excel_columns)

        # Connect to SQL server and run query
        cnxn = pyodbc.connect('Driver={SQL Server};'
                              'Server=' + self.server + ';'
                              'Database=' + self.db + ';'
                              'Trusted_Connection=yes;')

        query = SELECT  + ', '.join(sql_columns) +  FROM  + self.table
        query_result = pd.read_sql_query(query, cnxn)
        cnxn.close()

        # Compare data and return statement
        result = query_result[~query_result.isin(df)].dropna()
        if result.empty
            return All entries match the Excel file.
        else
            return Entries that do not match the Excel filen + str(result)

    def email_results(self, recipient_email, sender_email, password, excel_columns, sql_columns)
        # Compare data
        comparison_result = self.compare(excel_columns, sql_columns)

        # Create message
        message = MIMEText(comparison_result)
        message['Subject'] = 'Excel-SQL Comparison Results'
        message['From'] = sender_email
        message['To'] = recipient_email

        # Send email
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp.gmail.com, 465, context=context) as server
            server.login(sender_email, password)
            server.sendmail(sender_email, recipient_email, message.as_string())

if __name__ == __main__
    # Example usage
    excel_columns = [0, 1, 2] # columns to read from Excel file
    sql_columns = ['column1', 'column2', 'column3'] # columns to compare in the database table
    comparison = ExcelSQLComparison(table.xlsx, server_name, db_name, table_name)
    comparison.email_results(recipient@email.com, sender@email.com, sender_password, excel_columns, sql_columns)
