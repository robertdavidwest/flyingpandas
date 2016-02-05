__author__ = 'rwest'

import pandas as pd
import xlwings
import warnings

def _add_excel_formatting(sheet_name, spacing, column_formats, row_formats,
                          add_color_rows, autofit, include_index,
                          include_header):

    """Add number format and row colors to tables in excel report

    Parameters
    ----------
    data : pd.DataFrame
    sheetname : str
    spacing : dict of integers
        should contain 'start_row', 'start_col' and 'gap'
    column_formats : dict
    row_formats : dict
    add_color_rows : boolean
    autofit : boolean
    include_index : boolean
    include_header : boolean
    """
    startdatarow = spacing['startrow']
    if include_header:
        startdatarow += 1

    startdatacol = spacing['startcol']
    if include_index:
        startdatacol += 1

    # excel colors
    light_blue = (220, 230, 241)
    white = (255, 255, 255)

    # apply number formatting to all specified columns
    for col_num, format in column_formats.iteritems():
        if include_index:
            col_num += 1

        xlwings.Range(sheet_name,
                      (startdatarow, col_num),
                      (spacing['endrow'], col_num)
        ).number_format = format

    # apply number formatting to all specified rows
    for row_num, format in row_formats.iteritems():
        if include_header:
            row_num += 1

        xlwings.Range(sheet_name,
                      (row_num, startdatacol),
                      (row_num, spacing['endcol'])
        ).number_format = format

    if add_color_rows:
        # format color of rows in each table
        for i, row_num in enumerate(xrange(startdatarow,
                                           spacing['endrow'] + 1)):
            if i % 2 == 0:
                row_color = light_blue
            else:
                row_color = white

            xlwings.Range(sheet_name,
                          (row_num, startdatacol),
                          (row_num, spacing['endcol'])
            ).color = row_color

    # autofit
    if autofit:
        xlwings.Range(sheet_name,
                      (spacing['startrow'], spacing['startcol']),
                      (spacing['endrow'], spacing['endcol'])
        ).autofit()




class ExcelWriter(object):
    """
    Class for writing DataFrame objects into excel sheets, default is to use
    xlwt for xls, openpyxl for xlsx.  See DataFrame.to_excel for typical usage.

    Parameters
    ----------
    path : string
        Path to xls or xlsx file.
    engine : string (optional)
        Engine to use for writing. If None, defaults to
        ``io.excel.<extension>.writer``.  NOTE: can only be passed as a keyword
        argument.
    date_format : string, default None
        Format string for dates written into Excel files (e.g. 'YYYY-MM-DD')
    datetime_format : string, default None
        Format string for datetime objects written into Excel files
        (e.g. 'YYYY-MM-DD HH:MM:SS')
    """

    def __init__(self, path, engine=None, date_format=None,
                 datetime_format=None, **kwargs):

        self.pdwriter = pd.ExcelWriter(path=path, engine=engine,
                                     date_format=date_format,
                                     datetime_format=datetime_format,
                                     **kwargs)
        self._path = path
        self._sheet_name = []
        self._spacing = []
        self._column_formats = []
        self._row_formats = []
        self._add_color_rows = []
        self._autofit = []
        self._include_index = []
        self._include_header = []

    def close(self):
        """Save and close excel file and add specified formatting
        """
        self.pdwriter.close()

        # add excel formatting for each dataframe
        wb = xlwings.Workbook(self._path)
        for sheet_name, spacing, column_formats, row_formats, add_color_rows, \
            autofit, include_index, include_header in zip(self._sheet_name,
                                                          self._spacing,
                                                          self._column_formats,
                                                          self._row_formats,
                                                          self._add_color_rows,
                                                          self._autofit,
                                                          self._include_index,
                                                          self._include_header):

            _add_excel_formatting(sheet_name, spacing, column_formats,
                                  row_formats, add_color_rows, autofit,
                                  include_index, include_header)
        wb.save()
        wb.close()

    def to_excel(self, data, column_formats=None, row_formats=None,
                 row_format_col=None, add_color_rows=True, autofit=True,
                 sheet_name='Sheet1', na_rep='', float_format=None,
                 columns=None, header=True, index=True, index_label=None,
                 startrow=0, startcol=0, engine=None, merge_cells=True,
                 encoding=None, inf_rep='inf'):
        """
        Write DataFrame to a excel sheet using pandas.DataFrame.to_excel and
        store formatting preferences. Formats will be added to excel file after
        ``writer.close()``

        Parameters
        ----------
        data : pd.DataFrame
        column_formats : dict of string number formats
            e.g. {'col1' : '0.0%'}. Each number format is a string equivalent
            to excels custom number format
        row_formats : dict of string number formats
            e.g. {'row1' : '0.0%'}. Each number format is a string equivalent
            to excels custom number format. The string `row1` refers to the str
            variable row_format_col in the dataframe which must also be
            specified. e.g. if df['row_descriptions'] = ['thisrow', 'thatrow']
            then you could format the first row be specifiying:
              row_format_col='row_descriptions'
              row_formats = {'thisrow':  '0.0%'}
            row_formats takes priority over column_formats if both are specified
        row_format_col : str
            columns in data, used to specify row formatting
        add_color_rows : boolean (default True)
            color every other row of the dataframe being printed light blue
        autofit : boolean (default True)
            expand columns in dataframe to show all data
        sheet_name : string, default 'Sheet1'
            Name of sheet which will contain DataFrame
        na_rep : string, default ''
            Missing data representation
        float_format : string, default None
            Format string for floating point numbers
        columns : sequence, optional
            Columns to write
        header : boolean or list of string, default True
            Write out column names. If a list of string is given it is
            assumed to be aliases for the column names
        index : boolean, default True
            Write row names (index)
        index_label : string or sequence, default None
            Column label for index column(s) if desired. If None is given, and
            `header` and `index` are True, then the index names are used. A
            sequence should be given if the DataFrame uses MultiIndex.
        startrow : int
            upper left cell row to dump data frame
        startcol : int
            upper left cell column to dump data frame
        engine : string, default None
            write engine to use - you can also set this via the options
            ``io.excel.xlsx.writer``, ``io.excel.xls.writer``, and
            ``io.excel.xlsm.writer``.
        merge_cells : boolean, default True
            Write MultiIndex and Hierarchical Rows as merged cells.
        encoding: string, default None
            encoding of the resulting excel file. Only necessary for xlwt,
            other writers support unicode natively.
        cols : kwarg only alias of columns [deprecated]
        inf_rep : string, default 'inf'
            Representation for infinity (there is no native representation for
            infinity in Excel)

        >>> writer = ExcelWriter('output.xlsx')
        >>> writer.to_excel(df1,'Sheet1', column_formats={'Price': '$#,##0'})
        >>> writer.to_excel(df2,'Sheet2', column_formats={'Return': '0%'})
        >>> writer.close()

        """
        if not column_formats:
            column_formats = {}
        if columns:
            data = data.loc[:, columns]

        if row_formats and not row_format_col:
            err_msg = 'If using row_formats, you must also specify a row' \
                      'format column'
            raise ValueError(err_msg)

        if not row_formats:
            row_formats = {}

        # standard pandas to_excel
        data.to_excel(excel_writer=self.pdwriter, sheet_name=sheet_name,
                      na_rep=na_rep, float_format=float_format, columns=columns,
                      header=header, index=index, index_label=index_label,
                      startrow=startrow, startcol=startcol, engine=engine,
                      merge_cells=merge_cells, encoding=encoding, inf_rep=inf_rep)

        # convert row/col indexing from zero indexing to from one indexing for
        # excel
        startcol += 1
        startrow += 1

        # from here on out think indexing from 1!
        endrow = startrow + len(data)
        if not header:
            endrow -= 1

        endcol = startcol + len(data.columns)
        if not index:
            endcol -= 1

        spacing = {'startrow': startrow,
                   'startcol': startcol,
                   'endrow': endrow,
                   'endcol': endcol}

        # convert column names to column numbers in dataframe
        data = data.reset_index(drop=True)
        new_column_formats = {}
        for col_name, format in column_formats.iteritems():
            if col_name not in data.columns:
                warning_msg =  '"{}" is not a column name in dataframe, ' \
                            'no formatting will be applied for this colu' \
                            'mn'.format(col_name)
                warnings.warn(warning_msg)
                continue

            # get column number in dataframe
            col_num = 1
            for c in data.columns:
                if c == col_name:
                    break
                else:
                    col_num += 1

            # convert to column number in workbook
            col_num += spacing['startcol'] - 1
            new_column_formats[col_num] = format


        # convert row names to row numbers in dataframe
        if row_format_col:
            if row_format_col not in data.columns:
                err_msg = '"{}" is not a column name in dataframe' \
                          ''.format(row_format_col)
                raise KeyError(err_msg)

            row_names = data[row_format_col].drop_duplicates().tolist()

        new_row_formats = {}
        for row_name, format in row_formats.iteritems():
            if row_name not in row_names:
                warning_msg =  '"{}" is not a row name in the variable {} ' \
                           'no formatting will be applied for this row'.format(
                    row_name, row_format_col)
                warnings.warn(warning_msg)
                continue

            # get column number in dataframe
            row_num = 1
            for r in row_names:
                if r == row_name:
                    break
                else:
                    row_num += 1

            # convert to column number in workbook
            row_num += spacing['startrow'] - 1
            new_row_formats[row_num] = format


        # update formatting lists to apply formatting using xlwings after close
        self._spacing.append(spacing)
        self._column_formats.append(new_column_formats)
        self._row_formats.append(new_row_formats)
        self._add_color_rows.append(add_color_rows)
        self._autofit.append(autofit)
        self._sheet_name.append(sheet_name)
        self._include_index.append(index)
        self._include_header.append(header)
