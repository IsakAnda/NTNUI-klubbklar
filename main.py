'''
Main program for turning the answers to a excel file
'''
import pandas as pd
from person_class import Person
from converter import converter

answer_filename = 'svar_klubbklar_26_27.xlsx'   # Name of the answers from the form as a excel file.
order_filename = 'bestilling.xlsx'              # Name of the order filename
invoice_filename = 'faktura.xlsx'               # Name of the invoice filename
start_timestamp=pd.to_datetime('01/01/2025', format='%d/%m/%Y')
stop_timestamp = pd.to_datetime('01/01/2026', format='%d/%m/%Y')

def create_files(answer_filename, order_filename, invoice_filename, start_timestamp = None, stop_timestamp=None):
    df = pd.read_excel(answer_filename)
    df.fillna(0,inplace=True)

    order = []
    invoice = []
    people = []
    teams = {}

    for _, row in df.iterrows():
        answer = Person(row, start_timestamp, stop_timestamp)
        if answer.inside_timeframe:
            if answer.team not in teams:
                teams[answer.team] = [answer]
            else:
                teams[answer.team].append(answer)
            people.append(answer)
            order += answer.order
            invoice += answer.invoice

    if order != []:
        sheet_name = stop_timestamp.strftime('%d-%m-%Y')
        converter(pd.DataFrame(order), order_filename, sheet_name=sheet_name)         # Saves the order to 'bestilling.xlsx'
        converter(pd.DataFrame(invoice), invoice_filename, sheet_name=sheet_name)         # Saves invoice to "faktura.xlsx"

if __name__ == '__main__':
    create_files(answer_filename, order_filename, invoice_filename, start_timestamp, stop_timestamp)