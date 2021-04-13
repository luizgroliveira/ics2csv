import argparse
import csv
import glob
import logging  
import os
import vobject
from pathlib import Path, PosixPath, PurePath

logging.basicConfig(
    format='%(asctime)s - %(levelname)s - %(message)s', 
    datefmt='%m/%d/%Y %H:%M:%S', 
    level=logging.INFO)

class Expresso2office():
   def __init__(self, dir_csv="./output", dir_vcf="."):
      logging.debug(f"Passando aqui")
      self.dir_csv = PosixPath(dir_csv).expanduser()
      self.dir_vcf = PosixPath(dir_vcf).expanduser()
      self.format_date = "%m/%d/%Y"
      self.format_time = "%H:%M"
      self.format_date_time = f"{self.format_date} {self.format_time}"

   def checa_diretorios(self):
      logging.debug(f"Verificando diretorio de origem existe: {self.dir_vcf.resolve()}")
      if self.dir_vcf.is_dir():
         logging.info(f"Tratando arquivos vcf do diretorio: {self.dir_vcf.resolve()}")
      else:
         logging.info(f"Diretorio com arquivos vcf nao existe: {self.dir_vcf.resolve()}")
         exit(1)

      if not self.dir_csv.is_dir():
         logging.info(f"Diretorio para os arquivo csv nao existe, o mesmo sera criado: {self.dir_csv.resolve()}")
         try:
            self.dir_csv.mkdir()
         except OSError:
            logging.info(f"A criacao do diretorio {self.dir_csv.resolve()} falhou")
            exit(1)
         else:
            logging.log(f"Criou com sucesso o diretorio: {self.dir_csv.resolve()}")
    
   def convert2csv(self, file_vcf):
      try:
         data = Path(file_vcf).read_text()
         for cal in vobject.readComponents(data):
            file_csv = PurePath(self.dir_csv, f"address-{self.dir_vcf.name}.csv")
            logging.info("-" * 40)
            logging.info(f"Convertendo o arquivo VCF: {file_vcf}")
            logging.info(f"Arquivo cvs gerado: {file_csv}")
            with Path(file_csv).open(mode='a') as csv_out:
               csv_writer = csv.writer(csv_out, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
               if os.path.isfile(file_csv) and os.path.getsize(file_csv) == 0:
                  csv_writer.writerow(["firstName","middleName","lastName","company","jobTitle","workPhone","workPhone2","companyPhone","homePhone","homePhone2","mobilePhone","email","email2"])
               for line in cal.lines():
                  #import ipdb; ipdb.set_trace()
                  logging.debug(f"{file_vcf} name: {line.name}")
                  logging.debug(f"{file_vcf} behavior: {line.behavior}")
                  logging.debug(f"{file_vcf} encoded: {line.encoded}")
                  logging.debug(f"{file_vcf} group: {line.group}")
                  logging.debug(f"{file_vcf} params: {line.params}")
                  logging.debug(f"{file_vcf} serialize: {line.serialize()}")
                  logging.debug(f"{file_vcf} singletonparams: {line.singletonparams}")
                  logging.debug(f"{file_vcf} value: {line.value}")
                  logging.debug(f"{file_vcf} varlueRepr: {line.valueRepr()}")
                  if line.name == "VERSION":
                     try:
                        version = line.value
                     except:
                        version = ""
                  if line.name == "PRODID":
                     try:
                        prodid = line.value               
                     except:
                        prodid = ""
                  if line.name == "FN":
                     try:
                        full_name = line.value
                     except:
                        full_name = ""
                  if line.name == "N":
                     try:
                        first_name = line.serialize().split(":")[1].split(";")[1].lstrip()
                     except:
                        first_name = ""
                     try:
                        middle_name = line.serialize().split(":")[1].split(";")[2].lstrip()
                     except:
                        middle_name = ""
                     try:
                        last_name = line.serialize().split(":")[1].split(";")[0].lstrip()
                     except:
                        last_name = ""
                  if line.name == "UID":
                     try:
                        uid = line.value
                     except:
                        uid = ""
                  if line.name == "ORG":
                     try:
                        company = line.value[0]
                     except:
                        company = ""
                  if line.name == "TITLE":
                     try:
                        job_title = line.value
                     except:
                        job_title = ""
                  if line.name == "TEL":
                     if line.params['TYPE'] == ['WORK']:
                        try:
                           work_phone = line.value
                        except:
                           work_phone = ""
                     if line.params['TYPE'] == ['HOME']:
                        try:
                           home_phone = line.value
                        except:
                           home_phone = ""
                     if line.params['TYPE'] == ['CELL', 'WORK']:
                        try:
                           company_phone = line.value
                        except:
                           company_phone = ""
                     if line.params['TYPE'] == ['CELL', 'HOME']:
                        try:
                           mobile_phone = line.value
                        except:
                           mobile_phone = ""
                     if line.params['TYPE'] == ['FAX', 'WORK']:
                        try:
                           work_phone2 = line.value
                        except:
                           work_phone2 = ""
                     if line.params['TYPE'] == ['FAX', 'HOME']:
                        try:
                           home_phone2 = line.value
                        except:
                           home_phone2 = ""
                  if line.name == "EMAIL":
                     if line.params['TYPE'] == ["WORK"]:
                        try:
                           email = line.value
                        except:
                           email = ""
                     if line.params['TYPE'] == ["HOME"]:
                        try:
                           email2 = line.value
                        except:
                           email2 = ""
               csv_writer.writerow([first_name, middle_name, last_name, company, job_title, work_phone, work_phone2, company_phone, home_phone, home_phone2, mobile_phone, email, email2])
      except Exception as e:
         logging.fatal(f"Excessao nao mapeada: {e}")

   def realizar_parse(self):
      os.chdir(self.dir_vcf)
      for file_vcf in glob.glob("*.vcf"):
         self.convert2csv(file_vcf)


if __name__ == "__main__":
   parser = argparse.ArgumentParser()
   parser.add_argument("input", help="Diretorio de origem dos arquivos vcf")
   parser.add_argument("-o", "--output", help="Diretorio de saida do arquivo csv gerado (default: []./output)", default="./output")
   parser.add_argument("-v", "--verbose", help="Habilita o nivel de debug", action="store_true")
   args = parser.parse_args()

   if args.verbose:
      logging.info(f"Habilitando a execucao em modo debug")
      logging.getLogger().setLevel(logging.DEBUG)

   try:
      logging.info(f"Iniciando execucao")
      if args.output == "./output":
         args.output = Path(args.input).joinpath("output")

      logging.info(f"Diretorio de origem: {Path(args.input).resolve()}")
      logging.info(f"Diretorio de saida: {Path(args.output).resolve()}")
      logcvs = Expresso2office(dir_csv=Path(args.output).resolve(), dir_vcf=Path(args.input).resolve())
      logcvs.checa_diretorios()
      logcvs.realizar_parse()
      logging.info(f"Arquivos do diretorio processado com sucesso: {args.input}")
   except Exception as e:
      logging.error(f"Ocorreu erro ao executar a parse do arquivo {Path(args.input).resolve()}: {e}")
      os.sys.exit(1)