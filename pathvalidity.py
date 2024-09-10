from pathlib import Path
import os

p = Path('../Timesheet4Escalate/TimesheetTemplate')

if p.exists():
    print('Yes the path exists')

else:
    print('THe path doesnt exist')



print(os.getcwd())

