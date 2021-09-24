# Web-Report
Auto generate xlsx report from query result or 'list of dictionary' data

Return XLSX Report from the given 'List of Dictionary' data.
:param report_name: File name of the report before given datetime at the end of it
:param data: data with 'List of Dictionary' type, the key will be Header and the value will be inserted into each row 
:param start_date: For report title, if this param is filled, title will generated above header
:param end_date: For report title, start_date must be filled to show end_date. end date will shown after start_date
:param header: Enable/Disable header (title) option, header text will generated from data key's, underscores '_' will replaced with space and will capitalize unless all words is uppercase 
:param header_color: Color of the header, the default will be yellow (#FFFF00)
:param capitalize: Enable/Disable auto capitalize for header text
:param numbering: Enable/Disable Numbering. Number will generated at first column. The default will be True
:param auto_filter: Enable/Disable Auto Filter. 'Auto Filter' will filter first row. The default will be True, however.. this only active when params header is true
:param freeze_panes: Enable/Disable Freeze Pane. Freeze first column. The default will be True, however.. this only active when params header is true
:param freeze_panes_column: Freeze column from given integer (starting from 0). The default will be 0, however.. this only active when params header & freeze_panes is true
:param bottom_remark: Enable/Disable remark Give remark at the bottom of the workbook. The remark contains Downloader name & Download date
:return: xlsx file of the report generated
