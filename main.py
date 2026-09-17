'''
Main program for turning the answers to a excel file
'''

import pandas as pd
from datetime import datetime
from person_class import Person
from converter import converter

answer_filename = 'svar_klubbklar_26_27.xlsx'       # Name of the answers from the google form as a excel file.
order_filename = 'bestilling.xlsx'                  # Name of the order filename
invoice_filename = 'faktura.xlsx'                   # Name of the invoice filename


def create_files(answer_filename, order_filename, invoice_filename, start_timestamp = None, stop_timestamp=None):
    df = pd.read_excel(answer_filename)
    df.fillna(0,inplace=True)

    order = []
    invoice = []
    people = []
    teams = {}

    for index, row in df.iterrows():
        answer = Person(row, start_timestamp, stop_timestamp)
        if answer.inside_timeframe:
            if answer.team not in teams:
                teams[answer.team] = [answer]
            else:
                teams[answer.team].append(answer)
            people.append(answer)
            order += answer.order
            invoice += answer.invoice


    converter(pd.DataFrame(order), order_filename, sheet_name = f'{stop_timestamp[0]}.{stop_timestamp[1]}.{stop_timestamp[2]}')         # Saves the order to 'bestilling.xlsx'
    converter(pd.DataFrame(invoice), invoice_filename, sheet_name = f'{stop_timestamp[0]}.{stop_timestamp[1]}.{stop_timestamp[2]}')         # Saves invoice to "faktura.xlsx"

if __name__ == '__main__':
    create_files(answer_filename, order_filename, invoice_filename, start_timestamp=[1,1,2025], stop_timestamp=[15,9,2026])