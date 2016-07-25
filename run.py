import os, sys

days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]

# For free periods
for day in days:
	os.system("python main.py %s %s" % (day, sys.argv[1]))

# For daywise timetable printing
for day in days:
	os.system("python main.py %s %s" % (day, sys.argv[1]))
	os.system("wkhtmltopdf %s.html %s.pdf" % (day, day))
