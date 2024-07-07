import os
import openpyxl
import pandas as pd
import random
import configHandle



def main():

    # Construct the Excel file path
    excelValuePath = configHandle.currentExcelPath(__file__)

    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(excelValuePath, sheet_name="Receipt")

    # Use the first column as the key and the next 10 columns as values
    key_column = df.columns[0]
    value_columns = df.columns[1:]

    # Convert the Excel Header into List from df column.
    excelHeaders_list = df.columns.values.tolist()
    print("Excel Headers List:")
    print(excelHeaders_list)

    # Create a dictionary with key-value pairs
    key_value_dict = df.set_index(key_column)[value_columns].to_dict(orient="index")

    # Extract the values from the first column and convert it to a list
    alistOfKeyColumn = df.iloc[:, 0].tolist()

    print("\nList from the first column:")
    print(alistOfKeyColumn)

    # Randomly select a value from the list
    desired_key = random.choice(alistOfKeyColumn)
    print("Random desired value " + desired_key)

    # Example function call to retrieve a specific value
    desired_Value = str(configHandle.getDesiredKeyAndColumn(key_value_dict, "RE008", "AccountType"))

    print(desired_Value)

# def stepsToReadProductDetails():
#     print("stepsToReadProductDetails")
#     excelValuePath = configHandle.currentExcelPath(__file__)

    # Read the Excel file into a pandas DataFrame
    ReceiptEntry_AccountName_itemtablevalue_df = pd.read_excel(excelValuePath, sheet_name="AccountName")

    # Use the first column as the key and the next 10 columns as values
    key_itemtablevalue_column = ReceiptEntry_AccountName_itemtablevalue_df.columns[0]
    value_itemtablevalue_columns = ReceiptEntry_AccountName_itemtablevalue_df.columns[1:]

    # Create a dictionary with key-value pairs
    key_itemtablevalue_dict = ReceiptEntry_AccountName_itemtablevalue_df.set_index(key_itemtablevalue_column)[
        value_itemtablevalue_columns
    ].to_dict(orient="index")

    # Extract the values from the first column and convert it to a list
    alistOfKeyColumn_itemtablevalue = ReceiptEntry_AccountName_itemtablevalue_df.iloc[:, 0].tolist()

    print("List from the first itemtablevalue column:")
    print(alistOfKeyColumn_itemtablevalue)

    ConCodeValue = "RE008"

    aListOfItemIDFilteredWithSalesInvoiceID = configHandle.getlistItemIDWithSalesInvoiceID(
        ReceiptEntry_AccountName_itemtablevalue_df, ConCodeValue
    )
    print(f"List from itemtablevalue : {aListOfItemIDFilteredWithSalesInvoiceID}")

    # Read the Excel file into a pandas DataFrame
    ReceiptEntry_CC_itemtablevalue_df = pd.read_excel(excelValuePath, sheet_name="CC")

    # Use the first column as the key and the next 10 columns as values
    key_itemtablevalue_column = ReceiptEntry_CC_itemtablevalue_df.columns[0]
    value_itemtablevalue_columns = ReceiptEntry_CC_itemtablevalue_df.columns[1:]

    # Create a dictionary with key-value pairs
    key_itemtablevalue_dict = ReceiptEntry_CC_itemtablevalue_df.set_index(key_itemtablevalue_column)[
        value_itemtablevalue_columns
    ].to_dict(orient="index")

    # Extract the values from the first column and convert it to a list
    alistOfKeyColumn_itemtablevalue = ReceiptEntry_CC_itemtablevalue_df.iloc[:, 0].tolist()

    print("List from the first itemtablevalue column:")
    print(alistOfKeyColumn_itemtablevalue)

    ConCodeValue == "RE008"

    aListOfItemIDFilteredWithSalesInvoiceID = configHandle.getlist_ReceiptEntry_CC_ItemIDWithSalesInvoiceID(
        ReceiptEntry_CC_itemtablevalue_df, ConCodeValue
    )
    print(f"List from itemtablevalue : {aListOfItemIDFilteredWithSalesInvoiceID}")

    # Read the Excel file into a pandas DataFrame
    ReceiptEntry_RP_itemtablevalue_df = pd.read_excel(excelValuePath, sheet_name="RP")

    # Use the first column as the key and the next 10 columns as values
    key_itemtablevalue_column = ReceiptEntry_RP_itemtablevalue_df.columns[0]
    value_itemtablevalue_columns = ReceiptEntry_RP_itemtablevalue_df.columns[1:]

    # Create a dictionary with key-value pairs
    key_itemtablevalue_dict = ReceiptEntry_RP_itemtablevalue_df.set_index(key_itemtablevalue_column)[
        value_itemtablevalue_columns
    ].to_dict(orient="index")

    # Extract the values from the first column and convert it to a list
    alistOfKeyColumn_itemtablevalue = ReceiptEntry_RP_itemtablevalue_df.iloc[:, 0].tolist()

    print("List from the first itemtablevalue column:")
    print(alistOfKeyColumn_itemtablevalue)

    ConCodeValue == "RE008"

    aListOfItemIDFilteredWithSalesInvoiceID = configHandle.getlist_ReceiptEntry_RP_ItemIDWithSalesInvoiceID(
        ReceiptEntry_RP_itemtablevalue_df, ConCodeValue
    )
    print(f"List from itemtablevalue : {aListOfItemIDFilteredWithSalesInvoiceID}")


if __name__ == "__main__":
    main()