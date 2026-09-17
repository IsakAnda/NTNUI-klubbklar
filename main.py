'''
Main program for turning the answers to a excel file
'''

import pandas as pd
from person_class import Person
from converter import converter

answer_filename = 'svar_klubbklar_26_27.xlsx'       # Name of the answers from the google form as a excel file.
order_filename = 'bestilling.xlsx'                  # Name of the order filename
invoice_filename = 'faktura.xlsx'                   # Name of the invoice filename



df = pd.read_excel(answer_filename)
df.fillna(0,inplace=True)

order = []
invoice = []
people = []
teams = {}

for index, row in df.iterrows():
    answer = Person(row)
    if answer.team not in teams:
        teams[answer.team] = [answer]
    else:
        teams[answer.team].append(answer)
    people.append(answer)
    order += answer.order

for team in teams.values():
    team_total = 0
    for person in team:
        team_total+=person.total_price

    for person in team:
        for item in person.invoice:
            item['Totalt per lag'] = team_total
        invoice += person.invoice

converter(pd.DataFrame(order), order_filename)          # Saves the order to 'bestilling.xlsx'
converter(pd.DataFrame(invoice), invoice_filename)      # Saves invoice to "faktura.xlsx"
