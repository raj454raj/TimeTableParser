import xlrd
import sys

workbook = xlrd.open_workbook(sys.argv[2])
sheet = workbook.sheet_by_index(0)

main_dict = {}
daywise = {"Monday": {},
           "Tuesday": {},
           "Wednesday": {},
           "Thursday": {},
           "Friday": {},
           "Saturday": {}}

week_days = {"1": "Monday",
             "2": "Tuesday",
             "3": "Wednesday",
             "4": "Thursday",
             "5": "Friday",
             "6": "Saturday"}

def expand_days(days):
    """
        Examples:
        ---------
        1,2 => 1,2
        1-2 => 1,2
        1-4 => 1,2,3,4,5,6
        1,3-5 => 1,3,4,5
        1-3,6 => 1,2,3,6
        1-3,5-6 => 1,2,3,5,6
        5 => 5
    """

    final_string = ""
    parts = days.split(",")
    for part in parts:

        if final_string != "":
            final_string += ","

        tmp = part.split("-")
        if len(tmp) == 2:
            start = int(tmp[0])
            end = int(tmp[1])
            final_string += ",".join([str(x) for x in xrange(start, end + 1)])
        else:
            final_string += tmp[0]

    return final_string

if __name__ == "__main__":

    # To maintain the order of teachers
    ordered_teachers = []

    # Populate main_dict with sheet data
    for i in xrange(1, sheet.nrows):
        teacher_name = sheet.cell(i, 1).value
        ordered_teachers.append(teacher_name)
        main_dict[teacher_name] = {"designation": sheet.cell(i, 2).value}
        main_dict[teacher_name]["timetable"] = {"Monday": {},
                                                "Tuesday": {},
                                                "Wednesday": {},
                                                "Thursday": {},
                                                "Friday": {},
                                                "Saturday": {}}
        time_table = main_dict[teacher_name]["timetable"]
        for j in xrange(3, 11):
            period = j - 2
            content = sheet.cell(i, j).value.strip()
            content = content.split("\n")
            for line in content:
                line = line.strip()
                if line == "":
                    continue
                line = line.split(" ")
                days = expand_days(line[0]).split(",")
                for day in days:
                    final_print = line[1]
                    if len(line) > 2:
                        final_print += " " + line[2]
                    if daywise[week_days[day]].has_key(teacher_name):
                        daywise[week_days[day]][teacher_name][period] = final_print
                    else:
                        daywise[week_days[day]][teacher_name] = {}
                        daywise[week_days[day]][teacher_name][period] = final_print

                    time_table[week_days[day]][period] = line[1]

    if len(sys.argv) < 2:
        print "Pass day as argument"
        sys.exit(0)

    today = sys.argv[1]
    periods = {1: [], 2: [], 3: [], 4: [], 5: [], 6: [], 7: [], 8: []}
    for teacher in ordered_teachers:
        timetable = main_dict[teacher]["timetable"][today]
        for period in periods:
            if timetable.has_key(period) is False:
                periods[period].append((teacher, main_dict[teacher]["designation"]))
    prev = sys.stdout
    sys.stdout = open(today + ".txt", "w")

    for period in xrange(1, 9):
        print period
        print "======================================="
        for teacher in periods[period]:
            print teacher[0], "=>", teacher[1]
        print "\n\n"
    sys.stdout = prev
    prev = sys.stdout
    sys.stdout = open(today + ".html", "w")
    print """
          <html>
             <head>
                <style>
                    table {
                        font-size: 13px;
                        font-family: sans-serif;
                        border-collapse: collapse;
                    }
                    table, th, td {
                        border: 1px solid black;
                        padding: 4px;
                    }
                </style>
             </head>
          <body>
          <div style="text-align: center;">
            <h2>Kendriya Vidayalaya No.1 Sector 30, Gandhinagar</h3>
            <h3>TimeTable Session 2019-20 (%s)</h3>
            <p>As on 8th July, 2019</p>
          </div>
          <br />
          <table style='text-align: center; margin-left: auto; margin-right: auto;'>
                <tr>
                    <th>Sr No.</th>
                    <th>Teacher Name</th>
                    <th>1</th>
                    <th>2</th>
                    <th>3</th>
                    <th>4</th>
                    <th>5</th>
                    <th>6</th>
                    <th>7</th>
                    <th>8</th>
                </tr>
          """ % (today)
    cnt = 1
    for teacher in ordered_teachers:
        tr_string = ""
        tr_string += "\n\t<td>%d.</td>\n" % cnt
        cnt += 1
        tr_string += "\t<td style='text-align: left;'>" + teacher + " [" + main_dict[teacher]["designation"] + "]</td>\n"
        for period in xrange(1, 9):
            tr_string += "\t<td>"
            if teacher in daywise[today]:
                if daywise[today][teacher].has_key(period):
                    tr_string += daywise[today][teacher][period]
            tr_string += "</td>\n"
        print "<tr>%s</tr>" % tr_string
    print "</table>"
    print """
<br /><br /><br /><br /><br />
<div>
     <div style="float: left;">I/C TimeTable</div>
     <div style="float: right;">Principal</div>
</div>
          """
    print "</body></html>"
    sys.stdout = prev
