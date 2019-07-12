import common
import loadmail

# get list of all GitHub repos for task No 3
task3_repos = common.get_github_repo_names(common.settings.github_organization, 'os-task3')
# save all repos that are connected to AppVeyour before we any new repos are added
appveyour_repos = common.get_appveyor_project_repo_names()
# add repos to AppVeyor if they are not already there and trigger build for the newly added repos
added_repos = common.add_appveyor_projects_safely(list(task3_repos), trigger_build=True)
# show repos that were added to AppVeyour

list_mail = []
list_mail = loadmail.get_list_email()
loadmail.add_to_gsheets(list_mail)

for group in common.settings.os_groups:
    common.gsheet(group)
