import Functions as fun
import datetime

ms_id = '1xGENEoj0bLDss5uwQn0DdmAXhpZfyMDZXFzTESbyPx8'
week_end_str = raw_input('Enter Week End Date: ')

ts_lists = fun.get_spreadsheet_values(fun.get_timesheet_id(week_end_str,'0Bzzqy5DUeqkyZ1o5azktY3ZscnM'),"Invoice")
ms_lists = fun.get_spreadsheet_values(ms_id,"Define")

source = ['JPAZ-GD','JPBR-GD','JPCH-GD','JPFW-GD','JPLV-GD','JPVT-GD']


ts_lists[:] = [i for i in ts_lists if any(j in i for j in source)]

ms_lists[:] = [i for i in ms_lists if any(j in i for j in source)]

az = sum(x.count('JPAZ-GD') for x in ts_lists)
br = sum(x.count('JPBR-GD') for x in ts_lists)
ch = sum(x.count('JPCH-GD') for x in ts_lists)
fw = sum(x.count('JPFW-GD') for x in ts_lists)
lv = sum(x.count('JPLV-GD') for x in ts_lists)
vt = sum(x.count('JPVT-GD') for x in ts_lists)



we_date = fun.convert_to_date(week_end_str)



for i in ms_lists:
    start_date = fun.convert_to_date(str(i[4]))
    end_date = fun.convert_to_date(str(i[5]))
    if we_date >= start_date and we_date <= end_date:
        for j in ts_lists:
            if i[2] == j[2]:
                j[12] = i[9]
                j[13] = i[12]
                j[14] = str(i[17])[-6:]

az_lists = [i for i in ts_lists if source[0] in i]
br_lists = [i for i in ts_lists if source[1] in i]
ch_lists = [i for i in ts_lists if source[2] in i]
fw_lists = [i for i in ts_lists if source[3] in i]
lv_lists = [i for i in ts_lists if source[4] in i]
vt_lists = [i for i in ts_lists if source[5] in i]




if az > 0:
    print ('ERROR: Arizonia has Time no sheet created')
if br > 0:
    print ('ERROR: Brooklyn has Time no sheet created')
if ch > 0:
    print ('Chicago USCIS has Time')
    id = fun.copy_old_ss(fun.get_most_recent_file_id('0Bx6LO7D0JLPzQ3dWVzcxSmZtSXc'),week_end_str)
    invoice_num = fun.get_invoice_num(we_date,'0Bx6LO7D0JLPzQ3dWVzcxSmZtSXc')
    fun.write_to_spreadsheet(id,ch_lists,invoice_num)
if fw > 0:
    print ('ERROR: Fort Worth has Time no sheet created')
if lv > 0:
    print ('ERROR: Louisville has Time no sheet created')
if vt > 0:
    print ('Vermont has Time')
    id = fun.copy_old_ss(fun.get_most_recent_file_id('0Bx6LO7D0JLPzSHBkUTIwQnE0Vmc'),week_end_str)
    invoice_num = fun.get_invoice_num(we_date,'0Bx6LO7D0JLPzSHBkUTIwQnE0Vmc')

    fun.write_to_spreadsheet(id,vt_lists,invoice_num)
