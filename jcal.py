from __future__ import annotations
# PDF-Erzeugung
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, A3, letter, landscape
from reportlab.lib import colors
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
# Benutzeroberfläche
from tkinter import ttk, messagebox, filedialog as fd
import ttkbootstrap as ttkb
from ttkbootstrap.icons import Icon
import configparser

import icalendar     # .ics-Parser
import requests      # HTTP-Abruf

from pathlib import Path
import calendar
from icalendar import Calendar
import datetime as dt
from datetime import datetime, timedelta, date
from dateutil.relativedelta import relativedelta
import locale
import tempfile
import webbrowser
from collections import defaultdict

import sys, subprocess, os

# locale.setlocale(locale.LC_TIME, "de_DE.utf8")  # Deutsche Sprache setzen

## Helper-Funktionen
def get_ini_path():
   exe_dir = Path(sys.executable).parent
   try:
      with tempfile.TemporaryFile(dir=exe_dir):
         pass
      return exe_dir / 'jcal.ini'
   except Exception:
      return Path.home() / 'jcal' / 'jcal.ini'
   
def mm2pts(mm=0):
   return mm * 72 / 25.4

def pts2mm(pts=0):
   return pts *25.4 / 72

def open_file(path: str | os.PathLike):
   if sys.platform.startswith("win"):
      os.startfile(path)
   elif sys.platform == "darwin":
      subprocess.run(["open", path], check=False)
   else:
      subprocess.run(["xdg-open", path], check=False)

FONT_DIR = Path(__file__).with_name("fonts")
#INI_PATH = get_ini_path()
INI_PATH = Path.home() / "jcal.ini"
ICON_PATH = Path(__file__).with_name("icon") / "Chaninja-Chaninja-Folder-Program-Files.ico"

# ################################################################################################################
# ### Klasse zum Abrufen der Kalenderdaten und zur Erzeugung der PDF #############################################
# ################################################################################################################
class JCal:
   ## Klasseninitialisierung
   def __init__(self, header_prefix='Jahreskalender', fn='Jahreskalender.pdf', feedurl=''):
      self.startM = 1
      self.startY = int(dt.date.today().year)

      self.locale = {
         'monate': {
            'de': {
               'January':   'Januar',
               'February':  'Februar',
               'March':     'März',
               'April':     'April',
               'May':       'Mai',
               'June':      'Juni',
               'July':      'Juli',
               'August':    'August',
               'September': 'September',
               'October':   'Oktober',
               'November':  'November',
               'December':  'Dezember'
            }
         },
         'wochentage': {
            'de': {
               'Monday':    'Montag',
               'Tuesday':   'Dienstag',
               'Wednesday': 'Mittwoch',
               'Thursday':  'Donnerstag',
               'Friday':    'Freitag',
               'Saturday':  'Samstag',
               'Sunday':    'Sonntag'
            }
         },
         'wochentageK': {
            'de': {
               'Mon': 'Mo',
               'Tue': 'Di',
               'Wed': 'Mi',
               'Thu': 'Do',
               'Fri': 'Fr',
               'Sat': 'Sa',
               'Sun': 'So'
            }
         }
      }
      
   
   


   # ------------------------------------------------------------------------------------------------------------
   # Kalenderdaten sammeln und ordnen ---------------------------------------------------------------------------
   def parseEvents(self, start_month, start_year, feedurl):
      self.startM  = int(start_month)
      self.startY  = int(start_year)
      self.jahrgewechselt = False
      self.feedurl = feedurl
      
      # Kalenderdaten auslesen
      response = requests.get(feedurl)
      response.raise_for_status()
      cal = Calendar.from_ical(response.content)

      # Enddatum, geanu 1 Jahr später
      start_date = dt.datetime(self.startY, self.startM, 1)
      end_date   = start_date + relativedelta(years=1)

      # Dictionary zur Speicherung der Termine
      self.ebd = defaultdict(dict) # ebd = events_by_day
      self.fbm = defaultdict(dict) # fbm = footnotes by month

      # --- alle Termine durchlaufen und die relevanten Termine in den ------------------------------------------
      #     Dictionaries self.ebd und self.fbm speichern --------------------------------------------------------
      for element in cal.walk():
        
         if element.name == "VEVENT":
            
            dtstart = element.get('DTSTART').dt
            dtend   = element.get("DTEND")
            dtend = dtend.dt if dtend else None
            mehrtaegig_mZ = False
            mehrtaegig_oZ = False
            ganztaegig    = False
            
            ev_typ = "4-default"
            if dtend:
               if isinstance(dtstart, date) and not isinstance(dtstart, datetime): # ganztägig
                  dauer = (dtend - dtstart).days
                  if dauer > 1:                                                     # ganztägig mehrtägig
                     ev_typ = '1-mehrtaegig'
                  else:                                                             # ganztägig eintägig
                     ev_typ = '2-ganztaegig'
               else:                                                               # Ereignis mit Zeitangabe
                  dauer = dtend -dtstart
                  if dauer.days >= 1:                                               # mehrtägiges Ereignis mit Zeitangabe
                     ev_typ = '3-mehrtaegig_mZ'
                  else:                                                             # normaler Temrin mit Zeitangabe
                     pass
            else:
               pass                                                                # kein DTEND, Dauer kann nicht bestimmt werden

            evstart = dtstart if isinstance(dtstart, datetime) else datetime.combine(dtstart, datetime.min.time())
            if isinstance(dtend, datetime):
               evend = dtend
            elif dtend:
               evend = datetime.combine(dtend, datetime.min.time())
            else:
               evend = evstart + timedelta(days=1)
            
            # -- Lese den Termin nur ein, wenn er innerhalb des zu erfassenden Jahres liegt: start_date <=evstart <= end_date ODER evend >= start_date
            if (start_date.date() <= evstart.date() <= end_date.date()) or (evend.date() >= start_date.date()):
               print(f"Lese Termin ein.")
               summary       = element.get("SUMMARY")
               categories    = element.get('CATEGORIES')
               if categories:
                  categories_str = categories.to_ical().decode()
                  category_list  = [cat.strip() for cat in categories_str.split(',')]
               else:
                  categories_str = ''
                  category_list  = []
               
               print(f"Termin '{summary}': evstart = {evstart} - evend = {evend} | Typ: {ev_typ} | Kategorien: {categories_str}")
               
               ev_data = {
                  'kategorien': category_list,
                  'ev_start':   evstart,
                  'ev_end':     evend,
                  'summary':    summary,
                  'ev_typ':     ev_typ
               }

               # --- Mehrtagestermine müssen in das Dict. self.fbm (footnotes_by_month) für die fußnoten (Legende) des monats hinterlegt werden, 
               #     und auch die Termine, die die anzahl von maximal 4 anzeigbaren terminen pro zeile übersteigen - aber das sollte erst passieren, 
               #     nachdem die Tagestermien nach anfangszeit ortiert sind... kompliziert!! - kann erst nach dem kompletten Einlesen aller Termine erfolgen
               if ev_typ == "1-mehrtaegig" and not set(("Feiertag","Feiertage")) & set(category_list) and "Ferien" not in category_list: # -- ist ein mehrtägiger Termin, aber kein Feiertag oder Ferientermin
                  monate = self.betroffene_monate(evstart, evend)
                  #print(f'>> Der Termin erstreckt sich über {len(monate)} Monate, und zwar über {monate}')
                  for monat in monate:
                     key_fbm = monat
                     if key_fbm not in self.fbm: # --- falls es im dict self.fbm noch keinen eintrag für den aktuellen Monat gibt:
                        self.fbm[key_fbm] = []   #     Eintrag anlegen
                        #print(f'>> im dictionary fbm sind für den Monat [{key_fbm}] noch keine Fussnoten erfasst - Datenstaz self.fbm[{key_fbm}] wird angelegt.')
                     if not any(eintrag.get('fn_summary') == summary for eintrag in self.fbm[key_fbm]): # --- nur falls es in diesme Monat für diesen Summarytext noch keinen Eintrag gibt:
                        self.fbm[key_fbm].append({                                                      #     anlegen
                           'fn_evstart': evstart,
                           'fn_evend':   evend,
                           'fn_summary': summary,
                           'fn_typ':     ev_typ
                        })
                        #print(f">> im Array self.fbm['{key_fbm}'] gibt es noch keinen Eintrag mit der Summary '{summary}' - wird angelegt.")

               cur_day = evstart
               while cur_day < evend:
                  
                  cur_day_dt = cur_day.date()
                  # --- falls der tag im dict self.ebd (eventds_by_day) noch nicht angelegt wurde, dann anlegen mit Default-Daten
                  if cur_day_dt not in self.ebd:
                     self.ebd[cur_day_dt] = {
                        'termine':      [],
                        'is_feiertag':  False,
                        'is_ferientag': False,
                        'tagestexte':   []
                     }
                  
                  # --- jetzt die Daten für das Datum hinterlegen:
                  # --- 1. ggf. ändern, dass der tag ein Feiertag ist
                  if set(("Feiertag","Feiertage")) & set(category_list): self.ebd[cur_day_dt]['is_feiertag']  = True 
                  # --- 2. ggf. ändern, dass der tag ein Ferientag ist
                  if "Ferien" in category_list: self.ebd[cur_day_dt]['is_ferientag'] = True 
                  # ---NUR für den ersten Tag jeden Termins (die mehrtägigen Termine müssen nicht mehrfach gespeichert werden):
                  if cur_day == evstart or (cur_day == start_date and evstart < start_date):
                     # --- 3. Tagestexte hinzufügen: ein Array, dessen einträge später die zeilen der tagestexte sind: das wird aber erst nach dem Sortieren geschrieben
                     # self.ebd[cur_day_dt]['tagestexte'].append(summary)
                     # --- 4. die Termin-Daten an das array ...['termine'] anhängen
                     self.ebd[cur_day_dt]['termine'].append(ev_data)
                  elif ev_typ == '1-mehrtaegig' and not set(("Feiertag","Feiertage")) & set(category_list) and "Ferien" not in category_list: # bei mehrtägigen terminen, die über Monatsgrenzen hinweg gehen, auch am 1. der betroffenen Monate als Termin einfügen
                     if len(monate)>1:
                        print(f">> der Termin '{ev_data['summary']}' vom {ev_data['ev_start']} geht über mehr als einen Monat hinweg. Prüfe ob das Tuple ({cur_day.year}, {cur_day.month}) im array monate vorkommt (cur_day.day = {cur_day.day}): ", end=" -> ")
                        if (cur_day.year, cur_day.month) in monate and cur_day > evstart and cur_day.day==1:
                           print(f" ja (True)")
                           self.ebd[cur_day_dt]['termine'].append(ev_data)
                        else:
                           print("nein")

                  
                  
                  #print(f" der tag {cur_day_dt} erhält] den Status is_feiertag: {self.ebd[cur_day_dt]['is_feiertag']}")
                  print(f"...Der Tag {cur_day_dt} erhält den Status is_ferientag: {self.ebd[cur_day_dt]['is_ferientag']}")
                  cur_day += timedelta(days=1)

      # --- jetzt die Tagestexte und die mehrtägigen Termine (Fussnoten) sortieren ------------------------------    
      # --- 1-Tagestermine sortieren
      if len(self.ebd) > 0:
         for tag, daten in self.ebd.items():
            if 'termine' in daten:
               daten['termine'].sort(key=lambda ev: (ev['ev_typ'], ev['ev_start']))
            # --- Tagestexte neu schreiben (oder erst jetzt hier schreiben, muss vorher gar nicht erfolgen)
            if len(daten['tagestexte']) > 0: self.ebd[tag]['tagestexte'].clear()
            max_zeilen = 4
            ctr = 0
            for termin in daten['termine']:
               uhrzeit    = termin['ev_start'].strftime("%H:%M")
               zeilentext = termin['summary'].strip()
               if termin['ev_typ'] == '4-default': zeilentext = str(f"{uhrzeit} {zeilentext}")
               if ctr < max_zeilen: # --- bis zu 4 Tageseinträge? Diese als Tagestexte sortiert ausgeben
                  self.ebd[tag]['tagestexte'].append(zeilentext)
               else: # --- mehr als 4 Tageseinträge? Dann die letzten mit an die Fußnoten anhängen
                  key_fbm = (tag.year, tag.month)
                  if key_fbm not in self.fbm:
                     self.fbm[key_fbm] = []
                  self.fbm[key_fbm].append({
                     'fn_evstart': termin['ev_start'],
                     'fn_evend':   termin['ev_end'],
                     'fn_summary': zeilentext,
                     'fn_typ':     termin['ev_typ']
                  })
               ctr += 1
      # --- 2-Fußnoten sortieren nach Anfangszeit und dann sowohl als Legendeneinträge speichern als auch -------
      #     täglich für die Verweise in den Tageszeilen ---------------------------------------------------------
      print("\n<<<----------------------------------------------------------------------------------------------------------------------------->>>")
      print("<<< Sortiere die Fußnoten und erzeuge das dictionary self.fbd (footnotes_by_day). ---------------------------------------------->>>\n")
      self.fbd = defaultdict(dict) # --- fbd = footnotes_by_day, zur Speicherung und Ausgabe der Verweisnrn.
      print(f">> es gibt {len(self.fbm)} Elemente im Dictionary self.fbm (footnotes_by_month)")
      if len(self.fbm) > 0:
         for monat in self.fbm:
            #print(f"!! >>> Monat: {monat}: vor dem Sortieren: Anzahl der Elemente in self.fbm['{monat}']: {len(self.fbm[monat])}")
            self.fbm[monat].sort(key=lambda ev: ev['fn_evstart'].strftime("%y:%m:%D %H:%M"))
            #print(f"!! >>> Monat: {monat}: nach dem Sortieren: Anzahl der Elemente in self.fbm['{monat}']: {len(self.fbm[monat])}")
            #for i, v in enumerate(self.fbm[monat], start=1):
            #   print(f'Termin {i}: {v['fn_summary']} | fn_evstart: {v['fn_evstart']} | fn_evend: {v['fn_evend']}')

            for index, fn in enumerate(self.fbm[monat], start=1):
               erster_des_monats  = date(monat[0], monat[1], 1)
               letzter_des_monats = erster_des_monats + relativedelta(months=1) - relativedelta(days=1)
               vorletzter_des_termins = fn['fn_evend'].date() - relativedelta(days=1)
               # --- startdatum zum durchlaufen: Anfansdatum des termins oder erster_des_Monats, falls das Anfangsdatum in einem der vormonate liegt
               cur_date    = fn['fn_evstart'].date() if fn['fn_evstart'].date() >= erster_des_monats else erster_des_monats
               # --- Enddatum latest_date: abhängig von der Temrinart und Lage des letzten termintages --
               # --- Achtung: bei mehrtägigen Terminen leigt das enddatum am Folgetag um 00:00 Uhr!!
               if fn['fn_typ'] == "4-default":
                  latest_date = fn['fn_evend'].date() if fn['fn_evend'].date() <= letzter_des_monats else letzter_des_monats
               elif fn['fn_typ'] == "2-ganztaegig":
                  latest_date = fn['fn_evend'].date() if fn['fn_evend'].date() <= letzter_des_monats else letzter_des_monats
               elif fn['fn_typ'] == "1-mehrtaegig":
                   latest_date = vorletzter_des_termins if vorletzter_des_termins <= letzter_des_monats else letzter_des_monats
               else: 
                  latest_date = fn['fn_evend'].date() if fn['fn_evend'].date() <= letzter_des_monats else letzter_des_monats
               while (
                  #(fn['fn_typ'] == "1-mehrtaegig" and cur_date < latest_date) or
                  #(fn['fn_typ'] == "2-ganztaegig" and cur_date <= latest_date) or
                  #(fn['fn_typ'] == "4-default" and cur_date <= latest_date)
                  cur_date <= latest_date
               ):
                  #print(f'!! >>> Monat: {monat}: index: {index} | cur_date: {cur_date}')
                  if cur_date not in self.fbd:
                     #print(f"!! >>> Für das Datum '{cur_date} gibt es im Dictionary self.fbd noch keinen Eintrag -> wird angelegt und index: '{index}' hinzugefügt.")
                     self.fbd[cur_date] = []
                  self.fbd[cur_date].append(index)
                  cur_date += timedelta(days=1)
         




   def betroffene_monate(self, evstart:date, evend:date) -> list[tuple[int, int]]:
      monate = []
      aktuelles_datum = evstart.replace(day=1)
      end_datum       = evend
      #print(f"Funtkion betroffen_monate(evstart='{evstart}', evend='{evend}'): (evend - evstart).days = '{(evend - evstart).days}'")
      if (evend - evstart).days > 1: end_datum -= relativedelta(days=1)
      #print(f"Überprüfe aktuelles_datum ('{aktuelles_datum})' <= end_datum ({end_datum})")
      while aktuelles_datum <= end_datum:
         #print("...passt")
         monate.append((aktuelles_datum.year, aktuelles_datum.month))
         aktuelles_datum += relativedelta(months=1)
      return monate

   def is_ferientag(self, tag):
      tag = tag.date()
      if tag in self.ebd:
         return self.ebd[tag]["is_ferientag"]
      else:
         return False
   
   def is_feiertag(self, tag):
      tag = tag.date()
      if tag in self.ebd:
         return self.ebd[tag]["is_feiertag"]
      else:
         return False
   # ------------------------------------------------------------------------------------------------------------
   # PDF-Datei aus geordneten Kalenderdaten erstellen -----------------------------------------------------------
   def createPdf(self, fpath: Path, header: str):
      self.fpath  = fpath
      self.header = header.strip()
      # --- PDF-Konfiguration (Fonts, Fontgrößen, Dateiname, etc.....) ------------------------------------------
      self.pgsz = landscape(A3)
      self.wd, self.ht = self.pgsz
      # --- Seitenränder ----------------------------------------------------------------------------------------
      self.mgl  = mm2pts(8)                     # linker Rand
      self.mgr  = mm2pts(8)                     # rechter Rand
      self.mgb  = mm2pts(8)                     # unterer Rand
      self.mgt  = mm2pts(12)                    # oberer Rand
      self.mgm  = mm2pts(8)
      self.wdp  = self.wd - self.mgl -self.mgr  # Breite der Seite ohne Ränder, automatisch ermittelt
      self.htp  = self.ht - self.mgb - self.mgt # Höhe der Seite ohne Ränder, automatisch ermittelt
      # --- Fonts -----------------------------------------------------------------------------------------------
      pdfmetrics.registerFont(TTFont('Calibri',             FONT_DIR / "calibri.ttf"))
      pdfmetrics.registerFont(TTFont('Calibri-Kursiv',      FONT_DIR / "calibrii.ttf"))
      pdfmetrics.registerFont(TTFont('Calibri-Fett',        FONT_DIR / "calibrib.ttf"))
      pdfmetrics.registerFont(TTFont('Calibri-Fett-kursiv', FONT_DIR / "calibriz.ttf"))
      pdfmetrics.registerFont(TTFont('Calibri-Fein',        FONT_DIR / "calibril.ttf"))
      pdfmetrics.registerFont(TTFont('Calibri-Fein-Kursiv', FONT_DIR / "calibrili.ttf"))
      self.ftsz         = 10
      self.ftsz_termine =  6
      # self.ftfm         = 'Calibri'
      self.fontfms      = {'default': 'Calibri',
                           'header':  'Calibri-Fett',
                           'footer':  'Calibri-Kursiv'}
      self.fontsizes    = {'default':          10,
                           'header':           14,
                           'footer':           6,
                           'monatsname':       10,
                           'spalte_wtag':      12,
                           'spalte_termine':   6,
                           'spalte_fussnoten': 6,
                           'spalte_kw':        4,
                           'legende':          6}
      # --- Spaltenbereiten und Zeilenhöhen ---------------------------------------------------------------------
      self.anz_m = 6 # Anzahl der Monate pro Seite
      self.widths = {'Monat':         (self.wdp - self.mgm) / self.anz_m,
                     'Tag':           mm2pts(6),
                     'Wochentag':     mm2pts(6),
                     'Fussnote':      mm2pts(4),
                     'Kalenderwoche': mm2pts(3)}
      self.widths['Termintexte'] =    self.widths["Monat"] - (self.widths["Tag"] + self.widths["Wochentag"] + self.widths["Fussnote"] + self.widths["Kalenderwoche"])
      self.heights = {"Monat":        mm2pts(6),
                      "Tag":          mm2pts(8)}
      # --- PDF-Canvas initialisieren ---------------------------------------------------------------------------
      self.canv = canvas.Canvas(str(self.fpath), pagesize=self.pgsz)
      self.canv.setFont(self.fontfms['default'], self.fontsizes["default"])
      self.canv.setStrokeColor(colors.black)
      self.canv.setLineWidth(0.03)
      # --- Platzhalter im Header erstzen, z.B. mit Angaben zum Startjahr self.startY
      if self.startM == 1:
         str_hd_jahre = str(self.startY)
      else:
         str_hd_jahre = f"{self.startY}/" + str(self.startY+1)[2:4]
      self.header = self.header.format(jahre=str_hd_jahre)

      # --- PDF-Seiten hinzufügen--------------------------------------------------------------------------------
      self.pdf_addPage(1)
      self.pdf_addPage(2)
      # --- PDF-Datei speichern _--------------------------------------------------------------------------------
      self.canv.save()
   

   
   # --- PDF-Seite hinzufügen -----------------------------------------------------------------------------------
   def pdf_addPage(self, pgNo=1):
      self.canv.setLineWidth(0.03)
      self.pgNo=pgNo
      self.pgStartM = int(self.startM)
      self.pgStartY = int(self.startY)
      if self.jahrgewechselt:
         self.pgStartY += 1
      
      # --- äußere Rahmen zeichnen
      self.canv.rect(self.mgl, self.mgb, (self.wdp - self.mgm)/2, self.htp)
      self.canv.rect((self.wd+self.mgm)/2, self.mgb, (self.wdp-self.mgm)/2, self.htp)
      
      # --- Monate: Spaltenüberschrift
      mgm_offset = 0
      for mo in range(6):
         
         monatCtr = int(self.startM + mo)
         print(f"monatCtr: {monatCtr}")

         
         if pgNo==1:
            #''' Bestimme Monat und Jahr für die Seite 1'''
            #''' Bestimme Monat '''
            if monatCtr > 12:
               monat = monatCtr - 12
               if not self.jahrgewechselt:
                  self.jahrgewechselt = True
                  print(f"ggf. Jahreswechsel")
            else:
               monat = monatCtr
            #''' Bestimme Jahr '''
            if self.jahrgewechselt:
               jahr = int(self.pgStartY) + 1
            else:
               jahr = int(self.pgStartY)
         else:
            #''' Bestimme Monat und Jahr für die Seite 2'''
            #''' Bestimme Monat '''
            if monatCtr > 6:
               monat = monatCtr - 6
            else:
               monat = monatCtr + 6
            #''' Bestimme Jahr '''
            if mo == 0 and monat == 1:
                  self.jahrgewechselt = True
                  self.pgStartY = self.startY + 1
            if self.jahrgewechselt:
               jahr = int(self.startY) +1
            else:
               if monat < 6:
                  if not self.jahrgewechselt:
                     self.jahrgewechselt = True
                  jahr = int(self.startY) + 1
               else:
                  jahr = int(self.pgStartY)
         
         print(f"mo: {mo} - monat: {monat} - jahr: {jahr}")

         if(mo > 2): mgm_offset = self.mgm
         
         # --- Überschriften und Fußzeilen schreiben, nur alle 3 Monate, also auf jeder A4-Hälfte
         if mo==0 or mo==3:
            self.pdf_writeHeader(self.header, self.mgl + mo*self.widths["Monat"] + mgm_offset, self.ht - self.mgt +5) 
            self.pdf_writeFooter(mo, mgm_offset)
         
         # --- Monatsüberschriften
         try:
            monatsname = self.locale['monate']['de'][dt.date(jahr, monat, 1).strftime("%B")] + f" {str(jahr)[2:4]}"
         except Exception as exc:
            monatsname = dt.date(jahr, monat, 1).strftime("%B") + f" {str(jahr)[2:4]}"

         x  = self.mgl + mo*self.widths["Monat"] + mgm_offset
         y1 = self.ht - self.mgt - self.heights["Monat"]
         self.canv.setFillColorRGB(153/255, 204/255, 1)
         self.canv.rect(x, y1, self.widths["Monat"], self.heights["Monat"], 1, 1)
         self.canv.setFont(self.fontfms["default"], self.fontsizes["monatsname"])
         self.canv.setFillColor(colors.black)
         self.canv.drawCentredString(x + self.widths["Monat"]/2, y1+(self.heights["Monat"] -self.fontsizes["monatsname"])/2 + self.fontsizes["monatsname"]*0.2, monatsname)

         # --- Tage des Monats schreiben
         print(f"Gehe alle Tage des Monats {monatsname} durch")
         _, tage_des_monats = calendar.monthrange(jahr, monat)
         for tag in range(1, tage_des_monats +1):
            dt_tag = dt.datetime(jahr, monat, tag)
            try:
               wochentag = self.locale['wochentageK']['de'][dt_tag.strftime('%a')]
            except Exception as exc:
               wochentag = dt_tag.strftime('%a')
            y2 = self.ht - self.mgt - self.heights["Monat"] - (tag * self.heights["Tag"])
            self.canv.setFillColorRGB(1,1,1)

            # --- Einfärbung der Tageszeile + Tageszeile rahmen
            is_feiertag = self.is_feiertag(dt_tag)
            is_ferientag = self.is_ferientag(dt_tag)
            is_sonntag = (dt_tag.weekday() == 6)
            is_samstag = (dt_tag.weekday() == 5)
            #print(f">>> {dt_tag.date()}: ", end= "..")
            if is_feiertag or is_sonntag: # Sonn- und Feiertags
               #print(f" ist Sonntag ({is_sonntag}) oder is_feiertag ({is_feiertag})")
               self.canv.setFillColorRGB(1,224/255,104/255)
            elif is_samstag: # samstags
               #print(f" ist Samstag ({is_samstag})")
               self.canv.setFillColorRGB(1,235/255,153/255)
            elif is_ferientag:
               #print(f" is_ferientag ({is_ferientag})")
               self.canv.setFillColorRGB(218/255,1,163/255)
            else:
               pass
               #print(" ")
            self.canv.rect(x, y2, self.widths["Monat"], self.heights["Tag"], 1, 1)

            self.canv.setFillColor(colors.black)
            # --- Spalte Tagesnr.
            self.canv.setFont(self.fontfms['default'], self.fontsizes["spalte_wtag"])
            y_tage = y2 + (self.heights["Tag"] - self.fontsizes['spalte_wtag'])/2 + self.fontsizes["spalte_wtag"]*0.2
            self.canv.drawRightString(x + self.widths["Tag"] - 2.5, y_tage, str(tag))
            # --- Spalte Wochentag
            self.canv.drawCentredString(x + self.widths["Tag"] + (self.widths["Wochentag"]/2), y_tage, wochentag)
            # --- Spalte Kalenderwoche
            if dt_tag.weekday() == 0: # Kalenderwoche an jedem Montag anzeigen
               kw = dt_tag.isocalendar().week
               x_kw = x + self.widths["Monat"] - 2
               y_kw = y2 + (self.heights["Tag"] - self.fontsizes["spalte_kw"])/2 + self.fontsizes["spalte_kw"]*0.02
               self.canv.setFont(self.fontfms["default"], self.fontsizes["spalte_kw"])
               self.canv.drawRightString(x_kw, y_kw, str(kw))
            # --- Spalte Tagestexte
            x_termine = x + self.widths["Tag"] + self.widths["Wochentag"] + 2.5
            self.canv.setFont(self.fontfms["default"], self.fontsizes["spalte_termine"])
            y_termin = y2
            #print(f"dt_tag: {dt_tag} <=> ")
            if dt_tag.date() in self.ebd:
               #print("Ja, der Termin existiert in events_by_day")
               if 'tagestexte' in self.ebd[dt_tag.date()]:
                  ht_z  = self.fontsizes["spalte_termine"]-0.5
                  anz_t = len(self.ebd[dt_tag.date()]["tagestexte"])
                  ht_t  = anz_t * ht_z
                  y_termin = y2 + self.heights["Tag"]/2 + (anz_t-1)*ht_z/2 - ht_z/2 + 0.75
                  # self.ebd[dt_tag.date()]["tagestexte"].sort()
                  for termin in self.ebd[dt_tag.date()]["tagestexte"]:
                     self.canv.drawString(x_termine, y_termin, f"{termin}")
                     y_termin -= ht_z
            # --- Spalte Verweise auf Fußnoten in Legende
            if dt_tag.date() in self.fbd:
               if len(self.fbd[dt_tag.date()]) > 0:
                  ht_z  = self.fontsizes["spalte_fussnoten"]-0.5
                  anz_f = len(self.ebd[dt_tag.date()])
                  ht_f  = anz_f * ht_z
                  y_fn = y2 + self.heights["Tag"]/2# + (anz_f-1)*ht_z/2 - ht_z/2
                  x_fn = x + self.widths["Tag"] + self.widths["Wochentag"] + self.widths['Termintexte'] + self.widths["Fussnote"]/2
                  for fn in self.fbd[dt_tag.date()]:
                     self.canv.drawCentredString(x_fn, y_fn, f"{fn}")
                     y_fn -= ht_z
                     

            
         # --- fehlende Tageszeilen bis zu 31 auffüllen
         if tag < 31:
            rest_ht = (31-tag)*self.heights["Tag"]
            rest_y  = self.ht - self.mgt - self.heights["Monat"] - (31*self.heights["Tag"])
            self.canv.setFillColorRGB(0.81, 0.81, 0.81)
            self.canv.rect(x, rest_y, self.widths["Monat"], rest_ht, 1, 1)
         
         # --- Legende rahmen
         self.canv.setFillColorRGB(0,0,0)
         x_legende = x + 2
         y_legende = self.ht - self.mgt - self.heights["Monat"] - (31*self.heights["Tag"]) - mm2pts(2.5)
         self.canv.setFont(self.fontfms["default"], self.fontsizes["spalte_termine"])
         self.canv.drawString(x_legende, y_legende, f'Legende: ')
         ht_legende_rect = y_legende - self.mgb - mm2pts(1)
         self.canv.rect(x, self.mgb, self.widths["Monat"], ht_legende_rect, 1, 0)
         # --- Legende schreiben
         #str_fbm_mon = f'monat{dt_tag.date().month}'
         key_fbm = (dt_tag.date().year, dt_tag.date().month)
         if len(self.fbm[key_fbm]) > 0:
            ht_z = self.fontsizes["legende"]
            y_legendenzeile = y_legende - 2*ht_z
            for index, fn in enumerate(self.fbm[key_fbm], start=1):
               self.canv.drawString(x_legende, y_legendenzeile, f'{index}: {fn['fn_summary']}')
               y_legendenzeile -= ht_z

      
      
      # --- erzeugte Seite auf dem canvas-Objekt darstellen
      self.canv.showPage()

   # --- PDF-Datei speichern -------------------------------------------------------------------------------------
   def pdf_save(self):
      self.canv.save()
   # --- PDF-Kopf schreiben --------------------------------------------------------------------------------------
   def pdf_writeHeader(self, str_header='default_header', x_header=0, y_header=0, x_offset=10):
      self.canv.setFont(self.fontfms['header'], self.fontsizes["header"])
      self.canv.drawString(x_header + x_offset, y_header, str_header)
      self.canv.setFont(self.fontfms['default'], self.fontsizes["header"])
      self.canv.drawRightString(x_header + 3*self.widths["Monat"] - x_offset, y_header, f'{self.pgNo}. Halbjahr')
      self.canv.setFont(self.fontfms['default'], self.fontsizes["default"])
   
   # --- PDF-Fuß schreiben ---------------------------------------------------------------------------------------
   def pdf_writeFooter(self, mo=0, mgm_offset=0):
      jetzt = datetime.now().strftime("%d.%m.%Y - %H:%M:%S")
      y_footer = mm2pts(6)
      x_footer1 = self.mgl + mo*self.widths["Monat"] + mgm_offset
      x_footer2 = self.mgl + (mo+3)*self.widths["Monat"] + mgm_offset
      self.canv.setFont(self.fontfms['default'], self.fontsizes["footer"])
      self.canv.drawString(x_footer1, y_footer, f"Kalenderfeed: {self.feedurl[0:100]}")
      self.canv.drawRightString(x_footer2, y_footer, f"erstellt am: {jetzt}")
      self.canv.setFont(self.fontfms['default'], self.fontsizes["default"])

# ################################################################################################################
# ### GUI-Aufbau #################################################################################################
# ################################################################################################################
class App(ttkb.Window):
   def __init__(self):
      super().__init__(themename="darkly")
      self.title("Jahreskalender v0.1a: Kalenderstream -> PDF-Jahreskalender")
      #print(f'ICON_PATH = {ICON_PATH}')
      if ICON_PATH.exists():
         #print(f"ICON_PATH '{ICON_PATH}' exists")
         self.iconbitmap(ICON_PATH)

      wind_width    = 600
      wind_height   = 260
      wind_screenwd = self.winfo_screenwidth()
      wind_screenht = self.winfo_screenheight()
      wind_x        = (wind_screenwd // 2 ) - (wind_width // 2)
      wind_y        = (wind_screenht // 2 ) - (wind_height // 2)
      self.geometry(f"{wind_width}x{wind_height}+{wind_x}+{wind_y}")
      self.resizable(False, False)

      

      # --- Überschrift-Eingabe ----------------------------------------------------------------------------------
      cur_y = 20
      x1 = 20
      x2 = 155
      x3 = 460
      lbl_header = ttkb.Label(self, text="Kalender-Überschrift:")
      lbl_header.place(x=x1, y=cur_y)
      self.header_var = ttkb.StringVar()
      ent_header = ttkb.Entry(self, textvariable=self.header_var, width=62)
      ent_header.place(x=x2, y=cur_y, width=430)
      # --- URL-Eingabe ------------------------------------------------------------------------------------------
      cur_y = 60
      lbl_url = ttkb.Label(self, text="Kalender-URL (.ics):")
      lbl_url.place(x=x1, y=cur_y)
      self.url_var = ttkb.StringVar()
      ent_url = ttkb.Entry(self, textvariable=self.url_var, width=62)
      ent_url.place(x=x2, y=cur_y, width=x3-x2-10)
      btn_checkurl = ttkb.Button(self, text="Verbindung prüfen", command=self.on_check)
      btn_checkurl.place(x=x3, y=cur_y, width=125)
      # --- Statuslabel -------------------------------------------------------------------------------------------
      self.status = ttkb.Label(self, text="", foreground="grey")
      self.status.place(x=x2, y=cur_y + 30)
      # --- Startjahr --------------------------------------------------------------------------------------------
      ttkb.Label(self, text="Startjahr / Startmonat:").place(x=x1, y=140)
      self.year_var = ttkb.StringVar(value=str(dt.date.today().year))
      year_values = [str(y) for y in range(2000,2100)]
      cbx_year = ttkb.Combobox(self, textvariable=self.year_var, values=year_values, width=10)
      cbx_year.place(x=x2, y=140)
      self.update_idletasks()
      # --- Startmonat -------------------------------------------------------------------------------------------
      self.month_var = ttkb.StringVar(value=str(1))
      month_values = [str(m) for m in range(1,13)]
      cbx_month = ttkb.Combobox(self, textvariable=self.month_var, values=month_values, width=5)
      cbx_month.place(x=cbx_year.winfo_x() + cbx_year.winfo_width() + 10, y=140)
      # --- Buttons -----------------------------------------------------------------------------------------------
      self.btn_pdf = ttkb.Button(self, text="PDF-Jahreskalender erstellen", command=self.on_pdf, state="disabled")
      self.btn_pdf.place(x=20, y=190, width=560)
      # --- INI-Daten laden ---------------------------------------------------------------------------------------
      self.load_defaults()
      if self.url_var.get().strip(): # falls URL hinterlegt ist --> automatisch prüfen
         self.on_check()
   
   # --------------------------------------------------------------------------------------------------------------
   # --- Methode: Voreinstellungen (Profile) laden aus .ini-Datei -------------------------------------------------
   def load_defaults(self):
      cfg = configparser.ConfigParser()
      if INI_PATH.exists():
         cfg.read(INI_PATH, encoding="utf-8") # INI-Daten einlesen
      else:
         cfg["Settings"] = {}                 # leere INI-Daten anlegen
      
      sec = cfg.setdefault("Settings", {})
      # Standardwerte
      url    = sec.get("url",    "")
      year   = sec.get("year",   str(dt.date.today().year))
      month  = sec.get("month",  str(dt.date.today().month))
      header = sec.get("header", "Jahreskalender")
      # ins GUI übernehmen
      self.url_var.set(url)
      self.year_var.set(year)
      self.month_var.set(month)
      self.header_var.set(header)

      # INI ggf. neu schreiben, falls sie nicht da ist
      if not INI_PATH.exists():
         with INI_PATH.open("w", encoding="utf-8") as f:
            cfg.write(f)
   def save_defaults(self):
      cfg = configparser.ConfigParser()
      cfg["Settings"] = {'url':    self.url_var.get().strip(),
                         'year':   self.year_var.get().strip(),
                         'month':  self.month_var.get().strip(),
                         'header': self.header_var.get().strip()}
      with INI_PATH.open("w", encoding="utf-8") as f:
         cfg.write(f)

   
   # -------------------------------------------------
   # Kalenderstream prüfen
   def check_stream(self, url: str) -> tuple[bool, str | None]:
      url = url.strip()
      if url=="": return False, "Bitte geben Sie einen Link zum Kalenderstream ein."
      try:
         resp = requests.get(url, timeout=10)
         resp.raise_for_status()
         # sehr grobe Prüfung auf ics-Stream (kann je anch Quelle variieren)
         if b"BEGIN:VCALENDAR" not in resp.content[:200]:
            return False, "Das scheint kein gültiger iCalendar-Stream zu sein"
         return True, None
      except Exception as exc:
         return False, str(exc)

   # ----------------------------------------------------------------------------------------------------------
   # --- EVENT-Handler --------------------------------------------------------------------------------------------
   def on_check(self):
      url = self.url_var.get().strip()
      ok, statuserr = self.check_stream(url)
      if ok:
         self.status.config(text="Verbindung erfolgreich", foreground="green")
         self.btn_pdf.config(state="normal")
      else:
         self.status.config(text=f"{statuserr}", foreground="red")
         self.btn_pdf.config(state="disabled")

   def on_pdf(self):
      url = self.url_var.get().strip()
      self.save_defaults()
      try:
         jcal.parseEvents(self.month_var.get().strip(), self.year_var.get().strip(), url)
         with tempfile.TemporaryDirectory() as tmpdir:
            pdf_tmp = Path(tmpdir) / "jcal_tmp.pdf"
            jcal.createPdf(pdf_tmp, self.header_var.get().strip())
            
            # --- Speichern-unter-Dialog
            def_name = f"jahreskalender_{int(self.year_var.get())}_{int(self.month_var.get())}.pdf"
            file_path = fd.asksaveasfilename(
               parent=self,
               title="Jahreskalender-PDF speichern unter ...",
               defaultextension=".pdf",
               initialfile=def_name,
               filetypes=[("PDF-Datei", "*.pdf")])
            if not file_path:
               return
            Path(file_path).write_bytes(pdf_tmp.read_bytes())
            
         
         messagebox.showinfo("Fertig", f"PDF wurde erzeugt:\n{file_path}")
         open_file(file_path)
      except Exception as exc:
         messagebox.showerror("Fehler", f"PDF-Erzeugung schlug fehl:\n{exc}")

if __name__ == "__main__":
   jcal = JCal()
   App().mainloop()
   