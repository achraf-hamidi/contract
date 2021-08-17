import docx
# # import xlrd
# # import xlwt
# # from xlutils.copy import copy
# # from datetime import datetime
# # from dateutil.relativedelta import relativedelta
templ = 'CDD temp.docx'
doc = docx.Document(templ)

# # # pat = 'liste chauffeurs Transload 11. 2020§ (1).xlsx'

# # # # To open Workbook
# # # rb = xlrd.open_workbook(pat)
# # # wb = copy(rb)
# # # sheet = wb.get_sheet(0)
# # # sheet.write(3, 3, '22.03.2010')
# # # print(sheet.text)


# # DateTime myTime = new DateTime();

# # --Add 1 day
# # myTime.AddDays(1);

# # --Subtract 60 seconds
# myTime.AddSeconds(-60);

# --Add 2 years
# myTime.AddYears(2);


# import datetime

# today = datetime.date.today()
# deltt = datetime.timedelta(days=90)
# print(today)
# print(deltt)
# print(today+deltt)


i = 0
for x in doc.paragraphs[60].runs:

    print(i)
    print(x.text)
    i = i+1
