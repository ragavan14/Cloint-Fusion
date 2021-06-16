import ClointFusion as cf
import re
import pandas as pd
cf.launch_website_h("https://avinashtechlvr.github.io/ClointFusion-Training-Task-6/")
cf.scrape_save_contents_to_notepad(r'D:\ClointFusion\TASK-5')
cf.launch_any_exe_bat_application(r'D:\ClointFusion\TASK-5\notepad-contents.txt')
cf.scrape_save_contents_to_notepad(r'D:\ClointFusion\TASK-5\notepad-contents.txt')
cf.launch_any_exe_bat_application(r'D:\ClointFusion\TASK-5\New.xlsx')
cf.key_press('ctrl+V')

cf.key_press('ctrl+home')
cf.key_press('backspace')
cf.key_hit_enter()
cf.key_press('Up')

cf.key_press('alt+E')
cf.key_press('D')
cf.key_press('R')
cf.key_hit_enter()

cf.key_press('alt+E')
cf.key_press('D')
cf.key_press('R')
cf.key_hit_enter()
cf.key_press('ctrl+S')
cf.window_close_windows("Excel")
cf.window_close_windows("notepad")

row_column = cf.excel_get_row_column_count(excel_path=r"D:\ClointFusion\TASK-5\New.xlsx")

i = 1
while i < row_column[0]:
    one = cf.excel_get_single_cell(excel_path=r"D:\ClointFusion\TASK-5\New.xlsx", columnName="From", cellNumber=i)
    second = cf.excel_get_single_cell(excel_path=r"D:\ClointFusion\TASK-5\New.xlsx", columnName="To", cellNumber=i)
    third = cf.excel_get_single_cell(excel_path=r"D:\ClointFusion\TASK-5\New.xlsx", columnName="Amount", cellNumber=i)

    cf.launch_website_h("https://www.xe.com/currencyconverter/")
    cf.browser_mouse_click_h(
        cf.browser_locate_element_h('//*[@id="yie-close-button-80645131-85fd-57ea-a306-320c069f304e"]'))
    cf.browser_mouse_click_h(cf.browser_locate_element_h('//*[@id="midmarketFromCurrency"]'))
    cf.key_write_enter(one)
    cf.browser_mouse_click_h(cf.browser_locate_element_h('//*[@id="midmarketToCurrency"]'))
    cf.key_write_enter(second)
    cf.browser_mouse_click_h(
        cf.browser_locate_element_h('//*[@id="__next"]/div[2]/div[2]/section/div[2]/div/main/form/div[1]/div[1]'))
    cf.key_write_enter(str(third))
    amount = cf.browser_locate_element_h('//p[@class="result__BigRate-sc-1bsijpp-1 iGrAod"]',
                                         get_text=True)
    amount = re.findall("\d*\.?\d+", amount)

    if len(amount) > 1:

        str1 = ''.join(str(am) for am in amount)
        amount.clear()
        amount.append(str1)
    cf.excel_set_single_cell(excel_path=r"D:\ClointFusion\TASK-5\New.xlsx", columnName="Converted", cellNumber=i,
                             setText=amount[0])

    cf.browser_quit_h()
    i = i+2

df = pd.read_excel(r"D:\ClointFusion\TASK-5\New.xlsx", engine="openpyxl")
df = df[~df['From'].isnull()]
df.to_excel(r"D:\ClointFusion\TASK-5\New.xlsx", engine="openpyxl", index=False)
cf.launch_any_exe_bat_application(pathOfExeFile=r"D:\ClointFusion\TASK-5\New.xlsx")

cf.window_close_windows("Excel")

cf.launch_any_exe_bat_application('Outlook')
cf.key_press('alt+H')
cf.key_press('N')

cf.key_write_enter('asvaditya21@gmail.com')
cf.key_press('tab')
cf.key_press('tab')
cf.key_write_enter('Money Converter')

cf.key_press('alt+H')
cf.key_press('A+F')
cf.key_hit_enter()
