import os
import re
import pandas as pd
import matplotlib.pyplot as plt

TRANSACTIONS_DIRECTORY = "excel files/TD Transactions/Brokerage/"


def regex_parse_transaction(transaction):
    """Takes the parameter transaction which is a row of the dataframe containing the transactions from TD spreadsheet
    """
    patterns = {
        "Pattern_Bought_Sold": r"^(Bought|Sold)\s(\d)+\s(\w)+\s@\s(\d)+?(.(\d)+)$",
        "Pattern_Dividend": r"QUALIFIED DIVIDEND \((\w)+\)",
        "Pattern_Funding": r"CLIENT REQUESTED ELECTRONIC FUNDING RECEIPT \((FUNDS NOW)\)",
        "Pattern_Stock_Split": r"STOCK SPLIT \((\w)+\)",
        "Pattern_Security_Transfer": r"TRANSFER OF SECURITY OR OPTION IN \((\w)+\)",
    }

    for transaction_type, pattern in patterns.items():
        match_object = re.match(pattern, transaction['DESCRIPTION'])
        if match_object:
            print(match_object.group(0))


full_data_frame = pd.DataFrame()

for n in range(1, len(os.listdir(TRANSACTIONS_DIRECTORY)) + 1):
    transaction_file = pd.read_csv(TRANSACTIONS_DIRECTORY + f"transactions ({n}).csv")
    # Drop last row
    transaction_file.drop(transaction_file.tail(1).index, inplace=True)
    full_data_frame = pd.concat((full_data_frame, transaction_file), ignore_index=True)

# Gets all transactions:
for row in range(0, full_data_frame.shape[0]):
    trans = full_data_frame.loc[row, :]
    regex_parse_transaction(transaction=trans)
