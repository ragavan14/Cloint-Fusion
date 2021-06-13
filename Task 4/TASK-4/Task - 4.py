import ClointFusion as cf
cf.OFF_semi_automatic_mode()

row_column_count = cf.excel_get_row_column_count(r"Email.xlsx")
no_of_rows = row_column_count[0]
cf.browser_navigate_h("https://accounts.google.com/ServiceLogin/identifier?passive=1209600&continue=https%3A%2F%2Faccounts.google.com%2Fb%2F0%2FAddMailService&followup=https%3A%2F%2Faccounts.google.com%2Fb%2F0%2FAddMailService&flowName=GlifWebSignIn&flowEntry=ServiceLogin")
username = cf.gui_get_any_input_from_user("Enter your Mail id")
password = cf.gui_get_any_input_from_user("Enter your Password", password=True)


version = cf.excel_get_single_cell(excel_path=r"D:\ClointFusion\TASK-4\Email.xlsx", columnName="Hackathon version", cellNumber=0)
date = cf.excel_get_single_cell(excel_path=r"D:\ClointFusion\TASK-4\Email.xlsx", columnName="Date", cellNumber=0)
with open(file='Email.txt', mode='r') as fp:
    msg = fp.readlines()

fp.close()

keyword = 'Hackathon'
dateKeyword = 'Time:'
for i in range(len(msg)):
    line = msg[i]
    tempList = line.split()
    if keyword in tempList:
        index = tempList.index(keyword) + 1
        tempList[index] = str(version)
        msg[i] = ' '.join(tempList)

    if dateKeyword in tempList:
        index = tempList.index(dateKeyword)
        for k in range(3):
            tempList.pop(index + 1)
        tempList[index + 1] = str(date)
        msg[i] = ' '.join(tempList)

i = 1
for i in range(no_of_rows-1):
    cf.browser_wait_until_h("Compose")
    compose = cf.browser_locate_element_h('<div id=":kr" class="aic"><div class="z0"><div class="T-I T-I-KE L3" style="user-select: none" role="button" tabindex="0" jscontroller="eIu7Db" jsaction="click:dlrqf; clickmod:dlrqf" jslog="20510; u014N:cOuCgd,Kr2w4b" gh="cm">Compose</div></div></div>')
    cf.browser_mouse_click_h("compose")
    mail_id = cf.excel_get_single_cell(excel_path=r"D:\ClointFusion\TASK-4\Email.xlsx", columnName="Email Id", cellNumber=i)
    cf.browser_wait_until_h("To")
    cf.browser_write_h(Value=mail_id, User_Visible_Text_Element="To")
    cf.key_press("tab")
    cf.key_write_enter("Cloint Fusion Hackathon 9.0")
    cf.key_press("tab")

    # Printing in order
    for j in range(len(msg)):
        cf.key_write_enter(msg[j])
        if j < 19 or j > 24:
            cf.key_press("Backspace")
    cf.browser_wait_until_h("Folks")
    # Replacing the Folks with the name
    cf.search_highlight_tab_enter_open("Folks,", hitEnterKey="Yes", shift_tab="Yes")
    cf.key_press("backspace")
    name = cf.excel_get_single_cell(excel_path=r"D:\ClointFusion\TASK-4\Email.xlsx", columnName="Name", cellNumber=i)
    cf.key_write_enter(name)
    cf.key_press("backspace")
    cf.key_press("ctrl+enter")

