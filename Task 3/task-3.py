import ClointFusion as cf
cf.OFF_semi_automatic_mode()
cf.launch_any_exe_bat_application(r"D:\ClointFusion\TASK-3\excel.xlsx")
columns = cf.excel_get_all_header_columns(r"D:\ClointFusion\TASK-3\excel.xlsx")
cf.key_press("f5")
cf.key_write_enter("b1")
cf.key_press("ctrl + shift + l")
cf.key_press("alt + down + e")
cf.key_write_enter("Central")
cf.key_press("ctrl + a")
cf.key_press("ctrl + c")
cf.key_press("shift + f11")
cf.key_press("ctrl + v")
cf.key_press("ctrl+s")
cf.key_press("Alt+F4")
cf.excel_sort_columns(excel_path=r"D:\ClointFusion\TASK-3\excel.xlsx", sheet_name='Sheet2', firstColumnToBeSorted="Total")
cf.launch_any_exe_bat_application(r"D:\ClointFusion\TASK-3\excel.xlsx")
cf.key_press("right")
cf.key_press("left")
cf.key_press('Ctrl + Shift + right')
cf.key_press('alt')
cf.key_press('h')
cf.key_press('h')
cf.key_press('right')
cf.key_hit_enter()
cf.key_press('alt')
cf.key_press('h')
cf.key_press('f + c')
cf.key_press('down')
cf.key_press('right')
cf.key_press('right')
cf.key_press('right')
cf.key_press('right')
cf.key_press('right')
cf.key_press('right')
cf.key_press('right')
cf.key_hit_enter()
cf.key_press('right')
cf.key_press('right')
cf.key_press('right')
cf.key_press('right')
cf.key_press('right')
cf.key_press('right')
cf.key_press('Ctrl + Shift + down')
cf.key_press('alt')
cf.key_press('h')
cf.key_press('l')
cf.key_press('s')
cf.key_hit_enter()
username = cf.gui_get_any_input_from_user("outlook Username/Login")
password = cf.gui_get_any_input_from_user("your outlook Password",password=True)
cf.browser_navigate_h("outlook.live.com/mail/0/inbox")
cf.browser_mouse_click_h("Sign in")
cf.browser_write_h(username,"email")
cf.browser_mouse_click_h("next")
cf.browser_write_h(password,"password")
cf.browser_mouse_click_h("Sign in")

newmsg=cf.browser_locate_element_h('//*[@id="app"]/div/div[2]/div[2]/div[1]/div/div/div[1]/div[1]/div[2]/div/div/button/span')
cf.browser_mouse_click_h(newmsg)
cf.browser_wait_until_h("To")
cf.browser_write_h("ragavan.clointfusion@gmail.com",User_Visible_Text_Element="To")

cf.key_hit_enter()

cf.browser_mouse_click_h("Cc")
cf.browser_write_h("avinash.clointfusion@gmail.com", User_Visible_Text_Element="Cc")
cf.key_hit_enter()

cf.key_press('tab')

cf.key_write_enter("Task 4 Automation Test")

cf.key_press('tab')

cf.key_write_enter("This is a mail sent by a bot made of ClointFusion.Table Details:")
cf.key_hit_enter()

cf.launch_any_exe_bat_application(r"D:\ClointFusion\TASK-3\excel.xlsx")
cf.key_press("ctrl+a")
cf.key_press("ctrl+c")

cf.window_close_windows("excel")
cf.key_hit_enter()

cf.key_hit_enter()

cf.key_press("ctrl+v")
cf.key_hit_enter()

cf.key_write_enter("Thanks & Regards")
cf.key_write_enter("B.Ragavan")
cf.browser_mouse_click_h("Attach")
cf.browser_mouse_click_h("Browse this computer")
cf.key_write_enter(r'D:\ClointFusion\TASK-3\excel.xlsx')
cf.key_hit_enter()

cf.browser_mouse_click_h("Send")










