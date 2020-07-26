#!/usr/bin/python

# ics to csv example
# dependency: https://pypi.org/project/vobject/

import vobject
import csv
import os
import glob
from datetime import datetime

dir_csv="/tmp/csv/"
dir_ics="/tmp/ics"

if os.path.isdir(dir_ics):
    print "Tratando arquivos ics do diretorio: {}".format(dir_ics)
else:
    print "Diretorio com arquivos csv nao existe"
    exit(1)

if not os.path.isdir(dir_csv):
    print "Diretorio para os arquivo csv nao existe, o mesmo sera criado: ".format(dir_csv)
    try:
        os.mkdir(dir_csv)
    except OSError:
        print "A criacao do diretorio {} falhou".format(dir_csv)
    else:
        print "Criou com sucesso o diretorio: {}".format(dir_csv)

print "Salvando arquivos csv no diretorio {}".format(dir_csv)

    
def convert2csv(file_ics):
    print "Convertendo o arquivo: {}".format(file_ics)
    with open(dir_csv + '/' + file_ics + '.csv', mode='w') as csv_out:
        #csv_writer = csv.writer(csv_out, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        csv_writer = csv.writer(csv_out, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        csv_writer.writerow(["Subject", "StartDate", "StartTime", "EndDate", "EndTime", "Note","Attendees","Location","Organize"])
    
        # read the data from the file
        data = open(file_ics).read()
    
        # iterate through the contents
        for cal in vobject.readComponents(data):
            for component in cal.components():
                if component.name == "VEVENT":
                    # write to csv
                    #csv_writer.writerow([component.summary.valueRepr(),component.attendee.valueRepr(),component.dtstart.valueRepr(),component.dtend.valueRepr(),component.description.valueRepr()])
                    try:
                      summary = component.summary.valueRepr()
                    except:
                      summary = ""
    
                    try:
                      dtstart_date = component.dtstart.valueRepr().strftime("%m/%d/%Y")
                    except:
                      dtstart_date = ""
    
                    try:
                      dtstart_time = component.dtstart.valueRepr().strftime("%H:%M:%S")
                    except:
                      dtstart_time = ""
    
                    try:
                      dtend_date = component.dtend.valueRepr().strftime("%m/%d/%Y")
                    except:
                      dtend_date = ""
    
                    try:
                      dtend_time = component.dtend.valueRepr().strftime("%H:%M:%S")
                    except:
                      dtend_time = ""
    
                    try:
                      description = component.description.valueRepr()
                    except:
                      description = ""
    
                    #try:
                    #  attendee = component.attendee.valueRepr()
                    #except:
                    #  attendee = ""

                    try:
                        list_attendee = []
                        for attendee in component.attendee_list:
                            list_attendee.append(attendee.value.split(":")[1])
                    except:
                        list_attendee = []

                    try:
                        location = component.location.valueRepr()
                    except:
                        location  = ""

                    try:
                        organizer = component.organizer.valueRepr().split(":")[1]
                    except:
                        organizer = ""

                    try:
                        status = component.status.valueRepr()
                    except:
                        status = ""

                    #csv_writer.writerow([component.summary.valueRepr(),component.dtstart.valueRepr(),component.dtend.valueRepr()])
                    #print "date: {}, time: {}".format(dtstart_date, dtstart_time)
                    #csv_writer.writerow([summary, dtstart_date, dtstart_time, dtend_date, dtend_time, description.replace("\n","\\n"),attendee])
                    csv_writer.writerow([summary, dtstart_date, dtstart_time, dtend_date, dtend_time, description.replace("\n","\\n"),','.join(list_attendee), location, organizer, status])
    

os.chdir(dir_ics)
for file in glob.glob("*.ics"):
    convert2csv(file)
