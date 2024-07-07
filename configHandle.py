from configobj import ConfigObj
import os
import openpyxl
import pandas as pd
import random
import configHandle


def getitemIdColumnName(value):
    # sales Invoice
        configFilePath = "configTextReceiptEntry.txt"  # Read
        # getting the data of text file to store in config.
        config = ConfigObj(configFilePath)
        columnName = config[value]
        return columnName

def currentExcelPath(file):
    
    current_dir = os.path.dirname(os.path.abspath(__file__))
    excelValuePath = os.path.join(
        current_dir,
        "ReceiptEntryTemplate.xlsx",
    )
    return excelValuePath

def getlistItemIDWithSalesInvoiceID(ReceiptEntry_AccountName_itemtablevalue_df, salesinvoiceid):

    filtered_df = ReceiptEntry_AccountName_itemtablevalue_df[
        ReceiptEntry_AccountName_itemtablevalue_df[
            configHandle.getitemIdColumnName("Unique Id Column Name")
        ]
        == salesinvoiceid
    ]

    # Extract the itemid values into a list
    aListOfItemIDFilteredWithSalesInvoiceID = filtered_df[
        configHandle.getitemIdColumnName(
            "Receipt Details Account Name Unique Id Column Name"
        )
    ].tolist()
    return aListOfItemIDFilteredWithSalesInvoiceID

def getlist_ReceiptEntry_CC_ItemIDWithSalesInvoiceID(ReceiptEntry_CC_itemtablevalue_df, salesinvoiceid):

    filtered_df = ReceiptEntry_CC_itemtablevalue_df[
        ReceiptEntry_CC_itemtablevalue_df[
            configHandle.getitemIdColumnName("Unique Id Column Name")
        ]
        == salesinvoiceid
    ]

    # Extract the itemid values into a list
    aListOfItemIDFilteredWithSalesInvoiceID = filtered_df[
        configHandle.getitemIdColumnName(
            "Receipt Details CC Unique Id column Name"
        )
    ].tolist()
    return aListOfItemIDFilteredWithSalesInvoiceID

def getlist_ReceiptEntry_RP_ItemIDWithSalesInvoiceID(ReceiptEntry_RP_itemtablevalue_df, salesinvoiceid):

    filtered_df = ReceiptEntry_RP_itemtablevalue_df[
        ReceiptEntry_RP_itemtablevalue_df[
            configHandle.getitemIdColumnName("Unique Id Column Name")
        ]
        == salesinvoiceid
    ]

    # Extract the itemid values into a list
    aListOfItemIDFilteredWithSalesInvoiceID = filtered_df[
        configHandle.getitemIdColumnName(
            "Receipt Details RP Unique Id column Name"
        )
    ].tolist()
    return aListOfItemIDFilteredWithSalesInvoiceID

def getDesiredKeyAndColumn(key_value_dict, desired_key, desired_column):
    if desired_key in key_value_dict:
        # Check if the specified column exists in the values for the key
        if desired_column in key_value_dict[desired_key]:
            value_for_desired_column = key_value_dict[desired_key][desired_column]
            print(
                f"Value for key '{desired_key}' in column '{desired_column}': {value_for_desired_column}"
            )
            return value_for_desired_column
        else:
            print(f"Column '{desired_column}' not found for key '{desired_key}'.")
    else:
        print(f"Key '{desired_key}' not found in the dictionary.")