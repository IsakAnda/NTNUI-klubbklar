'''
Class for turning a pd.df to a nice NTNUI excel file.
'''

import pandas as pd
from pathlib import Path

from openpyxl.styles import Font, PatternFill


def converter(order: pd.DataFrame, filename, sheet_name = 'sheet1'):
    team_order = ['H-ELITE', 'H1', 'H2A', 'H2B', 'H2C', 'H3A', 'H3B', 'H3C','H4A', 'H4B', 'H4C', 'H4D', 'H-Bredde','D-ELITE', 'D1', 'D2A', 'D2B', 'D2C', 'D3A', 'D3B', 'D3C','D4A', 'D4B', 'D4C', 'D4D', 'D-Bredde']
    order = (
        order.assign(_lag_order=pd.Categorical(order['Lag'], categories=team_order, ordered=True))
        .sort_values(by=['_lag_order', 'Navn', 'Produkt'])
        .drop(columns='_lag_order')
    )

    file_path = Path(filename)

    teams = list(dict.fromkeys(order['Lag'].to_list()))
    color_list = ['background-color:#FFE599', 'background-color:#FFF2CC']
    def color_rows(row):
        return [color_list[teams.index(row.get('Lag'))%2]] * len(row)


    styled_df = order.style.apply(color_rows, axis=1)

    writer_kwargs = {'engine': 'openpyxl', 'mode': 'a' if file_path.is_file() else 'w'}
    if file_path.is_file():
        writer_kwargs['if_sheet_exists'] = 'replace'

    with pd.ExcelWriter(filename, **writer_kwargs) as writer:
        styled_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)

        worksheet = writer.sheets[sheet_name]

        green_fill = PatternFill(fill_type='solid', fgColor='93C47D')
        person_fills = [
            PatternFill(fill_type='solid', fgColor='FFE599'),
            PatternFill(fill_type='solid', fgColor='FFF2CC'),
        ]
        green_font = Font(bold=True)
        title_font = Font(bold=True, size=16)

        def merge_cells(start_row, start_column, end_row, end_column, value, fill, font=None):
            worksheet.merge_cells(
                start_row=start_row,
                start_column=start_column,
                end_row=end_row,
                end_column=end_column,
            )
            cell = worksheet.cell(start_row, start_column, value)
            for row in worksheet.iter_rows(
                min_row=start_row,
                max_row=end_row,
                min_col=start_column,
                max_col=end_column,
            ):
                for merged_cell in row:
                    merged_cell.fill = fill
                    if font:
                        merged_cell.font = font

        merge_cells(1, 1, 1, len(styled_df.columns), 'Bestillingsskjema NTNUI Volleyball', green_fill, title_font)
        for cell in worksheet[2]:
            cell.fill = green_fill
            cell.font = green_font

        last_team = order['Lag'].to_list()[0]
        last_row = 0
        for i in range(len(order['Lag'])):
            if order['Lag'].to_list()[i] != last_team:
                if last_row+2==i+1:
                    worksheet.cell(i + 2, 1, last_team).fill = green_fill
                    worksheet.cell(i + 2, 1).font = green_font
                else:
                    merge_cells(last_row + 3, 1, i + 2, 1, last_team, green_fill, green_font)
                if 'Totalt per lag' in order.columns:
                    merge_cells(last_row + 3, 8, i + 2, 8, f'=SUM(g{last_row+2}:g{i+1})', person_fills[teams.index(order['Lag'].to_list()[i-1])%2])
                last_team = order['Lag'].to_list()[i]
                last_row=i

        merge_cells(last_row + 3, 1, len(order['Lag']) + 2, 1, last_team, green_fill, green_font)
        if 'Totalt per lag' in order.columns:
            merge_cells(last_row + 3, 8, len(order['Lag']) + 2, 8, f'=SUM(g{last_row+2}:g{len(order["Lag"])+1})', person_fills[(len(teams)-1)%2])

        last_person = order['Navn'].to_list()[0]
        last_row = 0
        for i in range(len(order['Navn'])):
            if order['Navn'].to_list()[i] != last_person:
                person_fill = person_fills[teams.index(order['Lag'].to_list()[i-1])%2]
                merge_cells(last_row + 3, 2, i + 2, 2, last_person, person_fill)
                if 'Kostnad enkelt personer' in order.columns:
                    merge_cells(last_row + 3, 7, i + 2, 7, f'=SUM(f{last_row+2}:f{i+1})', person_fill)
                last_person = order['Navn'].to_list()[i]
                last_row=i

        person_fill = person_fills[(len(teams)-1)%2]
        merge_cells(last_row + 3, 2, len(order['Navn']) + 2, 2, last_person, person_fill)

        if 'Kostnad enkelt personer' in order.columns:
            merge_cells(last_row + 3, 7, len(order['Navn']) + 2, 7, f'=SUM(f{last_row+2}:f{len(order["Navn"])+1})', person_fill)