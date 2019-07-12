import common
import loadmail

list_mail = []
list_mail = loadmail.get_list_email()
loadmail.add_to_gsheets(list_mail)

for group in common.settings.os_groups:
    common.gsheet(group)
