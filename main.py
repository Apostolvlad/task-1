from openpyxl import load_workbook

i_write = 1 # с какой строки записывать в таблицу результата

i_start = 3 # номер столбца с которого идут данные для поиска
i_count = 5 # сколько всего столбцов

filename = 'table.xlsx'

def main():
    wb = load_workbook(filename)
    sheet_masters = wb['All_masters']
    sheet_districts = wb['districts']
    sheet_result = wb['output of inf for loading']

    masters = list(sheet_masters.rows)
    masters.pop(0)
    masters = tuple(filter(lambda x: not (x[0].value is None or x[0].value.startswith('rep')), masters))
    for row in list(sheet_districts.rows)[1:]:
        if row[0].value is None: break
        row_search = [row[i_col].value for i_col in range(i_start, i_start + i_count)]
        for row_master in masters: 
            if row_master[1].value in row_search:
                sheet_result[f'A{i_write}'] = f"INSERT INTO Address_District(addressId, districtId) values ((select addressId from Person_Address where personId = (select id from Person where beautify = '{row_master[0].value}') limit 1), {row[0].value});"
                i_write += 1

    wb.save(filename)

if __name__ == '__main__':
    main()