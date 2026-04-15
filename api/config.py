# Explicitly convert numeric and percentage columns after cleaning
numeric_columns = ["Last Price", "Last Price Change", "Current Value", "Today's Gain/Loss Dollar",
                   "Total Gain/Loss Dollar", "Cost Basis Total", "Average Cost Basis" ]  # Adjust as needed
percentage_columns = ["Today's Gain/Loss Percent", "Total Gain/Loss Percent", "Percent Of Account"]  # Add percentage columns here

starts_with_columns = ["The data and information", "Brokerage services are", "Date downloaded"]

account_columns = ["Account Number", "Account Name", "Symbol"]