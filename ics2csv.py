import vobject
import csv
import os
import glob
import sys
from pathlib import Path, PosixPath, PurePath

#dir_csv = "/home/lgro/git/ics2csv/temp/calendars/adm.office/result"
#dir_ics = "/home/lgro/git/ics2csv/temp/calendars/adm.office/"
format_date = "%m/%d/%Y"
format_time = "%H:%M"
format_date_time = f"{format_date} {format_time}"

if len(sys.argv) == 3:
    dir_ics = sys.argv[1]
    dir_csv = sys.argv[2]
else:
    print(f"Eh necessario dois parametros de entrada: dir_origem, dir_destino")

dir_csv = PosixPath(dir_csv).expanduser()
dir_ics = PosixPath(dir_ics).expanduser()

if dir_ics.is_dir():
   print(f"Tratando arquivos ics do diretorio: {dir_ics}")
else:
   print("Diretorio com arquivos ics nao existe")
   exit(1)

if not dir_csv.is_dir():
   print(f"Diretorio para os arquivo csv nao existe, o mesmo sera criado: {dir_csv}")
   try:
      dir_csv.mkdir()
   except OSError:
      print(f"A criacao do diretorio {dir_csv} falhou")
      exit(1)
   else:
      print(f"Criou com sucesso o diretorio: {dir_csv}")

print(f"Salvando arquivos csv no diretorio {dir_csv}")

    
def convert2csv(file_ics):
   try:
      data = Path(file_ics).read_text()
      for cal in vobject.readComponents(data):
         for component in cal.components():
            if component.name == "VEVENT":
               print("-" * 40)
               print(f"Convertendo o arquivo ICS VEVENT: {file_ics}")
               #file_csv = os.path.join(dir_csv, f"vevent-{os.path.basename(dir_ics)}.csv")
               file_csv = PurePath(dir_csv, f"vevent-{dir_ics.name}.csv")
               with Path(file_csv).open(mode='a') as csv_out:
                  csv_writer = csv.writer(csv_out, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                  #csv_writer.writerow(["Subject", "StartDate", "StartTime", "EndDate", "EndTime", "Note","Attendees","Location","Organize", "Status"])

                  if os.path.isfile(file_csv) and os.path.getsize(file_csv) == 0:
                     csv_writer.writerow(["Subject", "StartDate", "StartTime", "EndDate", "EndTime", "Location", "Attendees", "Note"])


                  # write to csv
                  #csv_writer.writerow([component.summary.valueRepr(),component.attendee.valueRepr(),component.dtstart.valueRepr(),component.dtend.valueRepr(),component.description.valueRepr()])
                  print(f"Arquivo de saida CSV VEVENT: {file_csv}")
                  try:
                    summary = component.summary.valueRepr()
                  except:
                    summary = ""

                  try:
                    dtstart_date = component.dtstart.valueRepr().strftime(format_date)
                  except:
                    dtstart_date = ""

                  try:
                    dtstart_time = component.dtstart.valueRepr().strftime(format_time)
                  except:
                    dtstart_time = ""

                  try:
                    dtend_date = component.dtend.valueRepr().strftime(format_date)
                  except:
                    dtend_date = ""

                  try:
                    dtend_time = component.dtend.valueRepr().strftime(format_time)
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

                  #csv_writer.writerow([summary, dtstart_date, dtstart_time, dtend_date, dtend_time, description.replace("\n"," "),','.join(list_attendee), location, organizer, status])
                  #csv_writer.writerow([summary, dtstart_date, dtstart_time, dtend_date, dtend_time, description.replace("\n"," "),','.join(list_attendee)])
                  csv_writer.writerow([summary, dtstart_date, dtstart_time, dtend_date, dtend_time, location, '|'.join(list_attendee), description.replace("\n"," ")])
            if component.name == "VTODO":
               #import ipdb; ipdb.set_trace()
               print("-" * 40)
               print(f"Convertendo o arquivo ICS VTODO: {file_ics}")
               file_csv = PurePath(dir_csv, f"vtodo-{dir_ics.name}.csv")
               with Path(file_csv).open(mode='a') as csv_out:
                  csv_writer = csv.writer(csv_out, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                  #csv_writer.writerow(["Subject", "StartDate", "StartTime", "EndDate", "EndTime", "Note","Attendees","Location","Organize", "Status"])
                  #csv_writer.writerow(["Subject", "StartDate", "StartTime", "EndDate", "EndTime", "Location", "Attendees", "Note"])
                  if Path(file_csv).is_file() and Path(file_csv).stat().st_size == 0:
                     #csv_writer.writerow(['class', 'created', 'dtstamp', 'dtstart', 'last-modified', 'organizer', 'priority', 'sequence', 'status', 'summary', 'uid', 'x-moz-lastack', 'x-tine20-container'])
                     csv_writer.writerow(['Subject', 'Body', 'DueDate', 'StartDate', 'Status', 'PercentComplete'])

                  print(f"Arquivo de saida CSV VTODO: {file_csv}")
                  try:
                    classe = component.getChildValue('class')
                  except:
                    classe = ""

                  try:
                     created = component.getChildValue('created')
                  except:
                     created = ""

                  try:
                     description = component.getChildValue('description')
                  except:
                     description = ""

                  try:
                     due = component.getChildValue('due').strftime(format_date_time)
                  except:
                     due = ""

                  try:
                     dtstamp = component.getChildValue('dtstamp')
                  except:
                     dtstamp = ""

                  try:
                     dtstart = component.getChildValue('dtstart').strftime(format_date_time)
                     #import ipdb; ipdb.set_trace()
                  except:
                     dtstart = "" 

                  try:
                     last_modified = component.getChildValue('last-modified')
                  except:
                     last_modified = ""

                  try:
                     organizer = component.getChildValue('organizer').split(":")[1]
                     email = component.organizer.params['EMAIL']
                     #import ipdb; ipdb.set_trace()
                  except:
                     organizer = ""

                  try:
                     percent_complete = component.getChildValue('percent_complete')
                  except:
                     percent_complete = ""

                  try:
                     priority = component.getChildValue('priority')
                  except:
                     priority = ""

                  try:
                     sequence = component.getChildValue('sequence')
                  except:
                     sequence = ""

                  try:
                     status_file = component.getChildValue('status')
                     status_dict = {
                        "NEEDS-ACTION": "NotStarted", # inclusive quando estiver em branco
                        "COMPLETED": "Completed",
                        "CONFIRMED": "Started",
                        "IN-PROCESS": "InProgress",
                        "CANCELLED": "Completed" # Tentar juntar no SUMMARY a palavra "Cancelada"
                     }
                     #import ipdb; ipdb.set_trace()
                     status = status_dict.get(status_file, "NotStarted")
                  except:
                     status = "NotStarted" 

                  try:
                     summary = component.getChildValue('summary')
                  except:
                     summary = ""

                  try:
                     uid = component.getChildValue('uid ')
                  except:
                     uid = "" 

                  try:
                     x_moz_lastack = component.getChildValue('x-moz-lastack')
                  except:
                     x_moz_lastack = ""

                  try:
                     x_tine20_container = component.getChildValue('x-tine20-container')
                  except:
                     x_tine20_container = "" 

                  #csv_writer.writerow([classe, created, dtstamp, dtstart, last_modified, organizer, priority, sequence, status, summary, uid, x_moz_lastack, x_tine20_container])
                  if status_file == "CANCELLED":
                     summary = f"[Cancelada] - {summary}"
                  csv_writer.writerow([summary, description, due, dtstart, status, percent_complete])
   except Exception as e:
       print(f"Excessao: {e}")


os.chdir(dir_ics)
for file in glob.glob("*.ics"):
    convert2csv(file)
#os.chdir(dir_ics)
#try:
#   for file in sorted(Path(dir_ics).glob(".ics")):
#      convert2csv(file)
#except Exception as e:
#   print(f"EXXXXXX: {e}")