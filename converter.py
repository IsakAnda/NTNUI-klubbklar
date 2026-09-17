'''
Class for turning a pd.df to a nice NTNUI excel file.
'''

import pandas as pd

def converter(order: pd.DataFrame, filename, sheet_name = 'sheet1'):
    team_order = ['H-ELITE', 'H1', 'H2A', 'H2B', 'H2C', 'H3A', 'H3B', 'H3C','H4A', 'H4B', 'H4C', 'H4D', 'H-Bredde','D-ELITE', 'D1', 'D2A', 'D2B', 'D2C', 'D3A', 'D3B', 'D3C','D4A', 'D4B', 'D4C', 'D4D', 'D-Bredde']
    order = (
        order.assign(_lag_order=pd.Categorical(order['Lag'], categories=team_order, ordered=True))
        .sort_values(by=['_lag_order', 'Navn', 'Produkt'])
        .drop(columns='_lag_order')
    )
    
    teams = list(dict.fromkeys(order['Lag'].to_list()))
    color_list = ['background-color:#FFE599', 'background-color:#FFF2CC']
    def color_rows(row):
        return [color_list[teams.index(row.get('Lag'))%2]] * len(row)


    styled_df = order.style.apply(color_rows, axis=1)

    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        styled_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)

        workbook  = writer.book
        worksheet = writer.sheets[sheet_name]

        green_format = workbook.add_format({'bg_color':'#93C47D', 'bold': True})
        title_format = workbook.add_format({'bg_color':'#93C47D', 'bold': True, 'size': 16})
        person_formats = [
            workbook.add_format({'bg_color': '#FFE599'}),
            workbook.add_format({'bg_color': '#FFF2CC'}),
        ]

        worksheet.merge_range(0,0,0, len(styled_df.columns)-1,'Bestillingsskjema NTNUI Volleyball', title_format)
        worksheet.write_row(1,0,styled_df.columns.to_list(), green_format)

        last_team = order['Lag'].to_list()[0]
        last_row = 0
        for i in range(len(order['Lag'])):
            if order['Lag'].to_list()[i] != last_team:
                if last_row+2==i+1:
                    worksheet.write(i+1,0, last_team, green_format)
                else:
                    worksheet.merge_range(last_row+2, 0, i+1, 0, last_team, green_format)
                if 'Totalt per lag' in order.columns:
                    worksheet.merge_range(last_row+2, 7, i+1, 7, f'=SUM(g{last_row+2}:g{i+1})', person_formats[teams.index(order['Lag'].to_list()[i-1])%2])
                last_team = order['Lag'].to_list()[i]
                last_row=i

        worksheet.merge_range(last_row+2, 0, len(order['Lag'])+1, 0, last_team, green_format)
        if 'Totalt per lag' in order.columns:
            worksheet.merge_range(last_row+2, 7, len(order['Lag'])+1, 7, f'=SUM(g{last_row+2}:g{len(order["Lag"])+1})', person_formats[(len(teams)-1)%2])        # FIXME writes =@SUMMER

        last_person = order['Navn'].to_list()[0]
        last_row = 0
        for i in range(len(order['Navn'])):
            if order['Navn'].to_list()[i] != last_person:
                worksheet.merge_range(last_row+2, 1, i+1, 1, last_person, person_formats[teams.index(order['Lag'].to_list()[i-1])%2])
                if 'Kostnad enkelt personer' in order.columns:
                    worksheet.merge_range(last_row+2, 6, i+1, 6, f'=SUM(f{last_row+2}:f{i+1})', person_formats[teams.index(order['Lag'].to_list()[i-1])%2])
                last_person = order['Navn'].to_list()[i]
                last_row=i

        worksheet.merge_range(last_row+2, 1, len(order['Navn'])+1, 1, last_person, person_formats[(len(teams)-1)%2])

        if 'Kostnad enkelt personer' in order.columns:
            worksheet.merge_range(last_row+2, 6, len(order['Navn'])+1, 6, f'=SUM(f{last_row+2}:f{len(order["Navn"])+1})', person_formats[(len(teams)-1)%2])