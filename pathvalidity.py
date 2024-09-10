from pathlib import Path
import os

p = Path('../Timesheet4Escalate/TimesheetTemplate.csv')

if p.exists():
    print('Yes the path exists')
    file = open('../Timesheet4Escalate/TimesheetTemplate.csv', 'r')
    content = file.read()
    print(content)
    

else:
    print('THe path doesnt exist')



print(os.getcwd())




op = open('../Timesheet4Escalate/TimesheetTemplate.csv')

