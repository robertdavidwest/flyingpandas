
import pandas
import flyingpandas

data = pandas.read_excel('example.xlsx')

# pandas to_excel :
writer = pandas.ExcelWriter('pd_example.xlsx')
data.to_excel(writer, index=False)
writer.close()

# flying pandas to_excel :
writer = flyingpandas.ExcelWriter('fp_example.xlsx')

column_formats = {'Production Budget': '$#,##0',
                  'Domestic Opening Weekend': '$#,##0',
                  'Domestic Box Office' : '$#,##0',
                  'Worldwide Box Office': '$#,##0'}
writer.to_excel(data, index=False, column_formats=column_formats,
                add_color_rows=True, autofit=True)
writer.close()

