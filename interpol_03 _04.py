"""
Created on Sat Jul 09 21:33:49 2022

@author: Stefan Klein

Hilfen unter https://www.delftstack.com/de/howto/python-pandas/
z.B. Sortieren

Vorgabewerte für andere Spindel hinzu gefügt, einfacher Plot hinzu gefügt, Formeln für Interpolation stimmen noch nicht.
In Excel-Vorlage: =@Moment(H9;H$3) ist VBA-Funktionsaufruf. In VBA-Modulen nachschauen.

Angefangen mit übernehmen von orig.Werten in sppedList,powerList,.......
powerList aus interpolation. Werte noch kontrollieren
torqueList erstellen aus Moment = Leistung * 9550 / Drehzahl geht erst mal. Werte noch kontrollieren
Übernahme von "Originalwerten aus Eingabe-Listen bei Leistung und Drehmoment. Ansonsten Interpolation/Berechnung. OK
Stromberechnung jetzt auch ok.
Berechnung S6 eingefügt

Synchron/Asynchron abfrage und kopieren von Strom und Drehmomentwerten aus erster Spalte in Drehzahl 0 erfolgt jetzt bei Eingabe der ersten Werte.

S6 in Diagramme
Diagramm Leistung/Drehmoment hinzu


aus strom-, leistung- und drehmomentlist dictionarys gemacht wg. kopieren in Listen für Graphen
Eingabe Komma in Punkt wandeln

Spannungswerte hinzugefügt

Wenn sich aus "Zwischendrehzahlen" für S1 und S6 unterschiedliche Anzahlen von Zwischenschritten ergeben, können keine Diagramme erstellt werden.
Lösung ->Drehzahllisten der Eingaben angleichen und in Solldrehzahllisten übernehmen. -> Erledigt


df muss Zeile/spalte tauschen sonst transform nicht möglich!

S6-Werte berechnen wenn nicht vorhanden, Angefangen in Zeile 262 bis 277
(Der folgende Code multipliziert beispielsweise jeden Wert in einem DataFrame mit drei mithilfe der Lambda-Funktion von Python:
DataFrame = DataFrame.transform(lambda y: y*3)
print(DataFrame))
Fehler in Zeile 262 : can't multiply sequence by non-int of type 'float'
https://www.delftstack.com/de/howto/python-pandas/apply-function-to-column-pandas/

https://www.delftstack.com/de/howto/python-pandas/pandas-convert-object-to-float/
!!!!  Mit konvertieren von Objekt in float geht es !!!!

Berechnung Smax-Werte hinzu und Graphen mit Smax

GraphenAuswahl hinzu
Vertikale Markierungen in Diagrammen bei Eingegebenen Drehzahlen hinzu

Speichern hinzu
Vorschaltdrossel in Abfrage eingefügt

Wegspeichern geändert. Geht noch nicht


todo:
Abfrage ob Abspeichern
Nenndrehzahl abfragen und Werte auf Nenndrehzahlwerte beschränken (Leistung und dann zugehörige Mommente und Ströme berechnen)




-----------------------------------------------------------------------------------------------
Endrehzahl auf nächsten vollen Tausender aufrunden wenn Enddrehzahl nicht glatter Tausender
"""
#import traceback
#import numpy as np
#import math as mt
import matplotlib.pyplot as plt
#from matplotlib.figure import Figure
import matplotlib.ticker as plticker
#from matplotlib import style
#from colorama import init
#from numpy.polynomial.polynomial import Polynomial
#from pandas.core.indexes.base import Index
#import matplotlib.ticker as plticker
import pandas as pd
import easygui as eg
import sys
from scipy.interpolate import interp1d
import os

import itertools 
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import datetime

#import scipy.interpolate as sci
#from sympy import interpolate


####################################################################
#now = datetime.datetime.now()
#today = now.strftime('%d.%m.%Y')

def eingabeS1():#(dictname):
    spinType = eg.choicebox(msg='Pick an item', title='', choices=['Synchron', 'Asynchron'], preselect=0)
    #print(spinType) #DEBUG
    if spinType == None:
        eg.msgbox(msg='Abbruch durch Benutzer!', title='Benutzerabbruch', ok_button='OK', image=None, root=None)
        sys.exit()
    eg.msgbox(msg='Die Enddrehzahlen bei den S1 und S6 Werten müssen gleich sein!\nDie Endrehzahl muss ganzzalig durch 100 teilbar sein! Bitte Sinnvoll auf oder abrunden.\n', title='Allgemeine Anweisungen', ok_button='OK', image=None, root=None)
    msg = "Geben sie vorhandene Daten ein"
    title = "Daten von Hersteller-MDB"
    fieldNames = ["Drehzahl_S1", 
                    "Spannung", 
                    "Strom", 
                    "Leistung", 
                    "Drehmoment"]

    fieldValues = ["","","","",""]  # Vorbelegung der Eingabefelder # soll verwendet werden

    fieldValues = eg.multenterbox(msg,title,fieldNames,fieldValues)
    for i in range(5): # Komma in Pkt. ändern
        #print(fieldValues[i]) #debug
        fieldValues[i] = fieldValues[i].replace(',','.')
        #print(fieldValues[i]) #debug
    if fieldValues == None:
        eg.msgbox(msg='Abbruch durch Benutzer!', title='Benutzerabbruch', ok_button='OK', image=None, root=None)
        sys.exit()
    dictnameS1 = dict(zip(fieldNames, fieldValues))
    #print(dictnameS1) #debug

        # Werte für Drehzahl 0 erzeugen
    if spinType == 'Synchron':
            # eingegebene Werte für Strom und Drehmoment in Spalte für Drehzahl 0 kopieren.
        dictnameS1_0 = {'Drehzahl_S1': '0', 'Spannung': '0', 'Strom': dictnameS1['Strom'], 'Leistung': '0', 'Drehmoment': dictnameS1['Drehmoment']}
        S1_DF = pd.DataFrame.from_dict(dictnameS1_0, orient="index", columns=[0])
            #   Eingegebene Werte in neue Spalte eintragen
        New_S1DF = pd.DataFrame.from_dict(dictnameS1, orient="index", columns=[S1_DF.shape[1]])
        S1_DF = pd.concat([S1_DF, New_S1DF], axis=1)
    else:
            #   bei Asynchron erste Spalte mit allen Werten = 0 eintragen
        dictnameS1_0 = {'Drehzahl_S1': '0', 'Spannung': '0', 'Strom': '0', 'Leistung': '0', 'Drehmoment': '0'}
        S1_DF = pd.DataFrame.from_dict(dictnameS1_0, orient="index", columns=[0])
            #   Eingegebene Werte in neue Spalte eintragen
        New_S1DF = pd.DataFrame.from_dict(dictnameS1, orient="index", columns=[S1_DF.shape[1]])
        S1_DF = pd.concat([S1_DF, New_S1DF], axis=1)
    
    #print("Primer",S1_DF) #debug
    #print("\n") #debug    
    ####################################################

    while True:
        frageS1 = eg.buttonbox(msg="weitere Eingaben S1?", choices = ("No", "Yes") )
        if frageS1 == "Yes":
            msg = "Geben sie vorhandene Daten ein"
            title = "Daten von Hersteller-MDB"
            fieldNames1 = ["Drehzahl_S1", 
                    "Spannung", 
                    "Strom", 
                    "Leistung", 
                    "Drehmoment"]
            WerteS1 = []
            WerteS1 = eg.multenterbox(msg,title,fieldNames1,WerteS1)
            for i in range(5): # Komma in Pkt. ändern
                #print(WerteS1[i]) #debug
                WerteS1[i] = WerteS1[i].replace(',','.')
                #print(WerteS1[i]) #debug
            if WerteS1 == None:
                eg.msgbox(msg='Abbruch durch Benutzer!', title='Benutzerabbruch', ok_button='OK', image=None, root=None)
                sys.exit()
            dictnameS1_1 = dict(zip(fieldNames1, WerteS1))
            New_S1DF = pd.DataFrame.from_dict(dictnameS1_1, orient="index", columns=[S1_DF.shape[1]])
            S1_DF = pd.concat([S1_DF, New_S1DF], axis=1)
            #print("Neuevo", S1_DF)
            #print("Primer",S1_DF) #debug
            #print("\n") #debug
        else:
            break
    return (S1_DF, spinType)

def eingabeS6(spinType):#(dictname):
    eg.msgbox(msg='Die Enddrehzahlen bei den S1 und S6 Werten müssen gleich sein!\nDie Endrehzahl muss ganzzalig durch 100 teilbar sein! Bitte Sinnvoll auf oder abrunden.\n', title='Allgemeine Anweisungen', ok_button='OK', image=None, root=None)
    msg = "Geben sie vorhandene Daten ein"
    title = "Daten von Hersteller-MDB"
    fieldNames = ["Drehzahl_S6", 
                    "Spannung_S6", 
                    "Strom_S6", 
                    "Leistung_S6", 
                    "Drehmoment_S6"]

    fieldValues = ["","","","","","",""]  # Vorbelegung der Eingabefelder # soll verwendet werden
    fieldValues = eg.multenterbox(msg,title,fieldNames,fieldValues)
    for i in range(5): # Komma in Pkt. ändern
        #print(fieldValues[i]) #debug
        fieldValues[i] = fieldValues[i].replace(',','.')
        #print(fieldValues[i]) #debug
    dictnameS6 = dict(zip(fieldNames, fieldValues))
####################################################
    if spinType == 'Synchron':
            # eingegebene Werte für Strom und Drehmoment in Spalte für Drehzahl 0 kopieren.
        dictnameS6_0 = {'Drehzahl_S6': '0', 'Spannung_S6': '0', 'Strom_S6': dictnameS6['Strom_S6'], 'Leistung_S6': '0', 'Drehmoment_S6': dictnameS6['Drehmoment_S6']}
        S6_DF = pd.DataFrame.from_dict(dictnameS6_0, orient="index", columns=[0])
            #   Eingegebene Werte in neue Spalte eintragen
        New_S6DF = pd.DataFrame.from_dict(dictnameS6, orient="index", columns=[S6_DF.shape[1]])
        S6_DF = pd.concat([S6_DF, New_S6DF], axis=1)
    else:
            #   bei Asynchron erste Spalte mit allen Werten = 0 eintragen
        dictnameS6_0 = {'Drehzahl_S6': '0', 'Spannung_S6': '0', 'Strom_S6': '0', 'Leistung_S6': '0', 'Drehmoment_S6': '0'}
        S6_DF = pd.DataFrame.from_dict(dictnameS6_0, orient="index", columns=[0])
        #   Eingegebene Werte in neue Spalte eintragen
        New_S6DF = pd.DataFrame.from_dict(dictnameS6, orient="index", columns=[S6_DF.shape[1]])
        S6_DF = pd.concat([S6_DF, New_S6DF], axis=1)
    
    #print("Primer",S6_DF) #debug
    #print("\n") #debug    

    #S6_DF = pd.DataFrame.from_dict(dictnameS6, orient="index")
    while True:
        frageS6 = eg.buttonbox(msg="weitere Eingaben S6?", choices = ("No", "Yes"))
        if frageS6 == "Yes":
            msg = "Geben sie vorhandene Daten ein"
            title = "Daten von Hersteller-MDB"
            fieldNames1 = ["Drehzahl_S6", 
                    "Spannung_S6", 
                    "Strom_S6", 
                    "Leistung_S6", 
                    "Drehmoment_S6"]
            WerteS6 = []
            WerteS6 = eg.multenterbox(msg,title,fieldNames1,WerteS6)
            for i in range(5): # Komma in Pkt. ändern
                #print(WerteS6[i]) #debug
                WerteS6[i] = WerteS6[i].replace(',','.')
                #print(WerteS6[i]) #debug
            dictnameS6_1 = dict(zip(fieldNames1, WerteS6))
            New_S6DF = pd.DataFrame.from_dict(dictnameS6_1, orient="index", columns=[S6_DF.shape[1]])
            S6_DF = pd.concat([S6_DF, New_S6DF], axis=1)
            
        else:
            break
    return S6_DF
        # Funktionen zum Umrechnen der S1-Werte in S6-Werte und Smax-Werte
def mult1_3(x):
    #print(x) #debug
    x = x * 1.3
    #print(x) #debug
    return x

def mult1_4(x):
    #print(x) #debug
    x = x * 1.4
    #print(x) #debug
    return x

def save_file_dialogs(default_filename, savepath, extension):
    filename = savepath + "\\AS" + default_filename + "." + extension
    while 1:
        #filename = savepath + default_filename + "." + extension
        if os.path.exists(filename):
            ok_to_overwrite = eg.buttonbox(msg="Datei %s besteht bereits. Überschreiben?" %(filename), choices = ("No", "Yes") )
            if ok_to_overwrite == "Yes":
                return filename
                break
            elif ok_to_overwrite == "No":
                filename = eg.filesavebox(msg = "Dateiname eingeben (die Erweiterung %s wird automatisch angehängt)?\n --> Programm abbrechen mit 'Abbrechen'" %(extension), default = filename)
                if filename is None:
                    sys.exit()
                    #return None
                else:
                    continue
        else:
            return filename

def make_plot(
                InfoValues, 
                Drehzahl, 
                drehzahllist, 
                speedList, 
                currentList, 
                currentList6, 
                currentListMax, 
                powerList, 
                powerList6, 
                powerListMax, 
                voltageList, 
                voltageList6, 
                voltageListMax, 
                torqueList,
                torqueList6,
                torqueListMax,
                spinType,
                graphen
                ):
    ######################################  Diagramme  ############################################################################
        # Diagramme erstellen
    TitelInfo = (InfoValues[0] + '   AS' + InfoValues[1] + '   ' + spinType + '   ' + InfoValues[2] + '   ' + InfoValues[3])

    xaxs_max = Drehzahl.max()
    xaxs_min = 0

    cm = 1/2.54  # centimeters in inches

    fig0 = plt.figure(figsize=(26*cm, 13*cm))
    ax = fig0.add_subplot(111)
    ax.plot(speedList, currentList, label = 'Current_S1', color ='black')
    if 'S6' in graphen:
        ax.plot(speedList, currentList6, label = 'Current_S6 (40%, tS=2min)', linestyle = 'dashdot', color ='black')
    if 'Max-Werte' in graphen:
        ax.plot(speedList, currentListMax, label = 'Current_Max ', linestyle = 'dotted', color ='black')
    
    ax.set_xlabel('Speed [1/min]')
    ax.set_ylabel("Current [A]")
    ax.set_xlim(xaxs_min,xaxs_max)
    ax.set_ylim(0,None)
    #ax.yaxis.set_major_locator(plticker.MultipleLocator(base=10))

    plt.suptitle(TitelInfo, fontsize=14, y =1.0)

    ax.grid(b=True, which='major', color='#666666', linestyle='-')#Gitternetz Show the major grid lines with dark grey lines
    ax.minorticks_on()# Show the minor grid lines with very faint and almost transparent grey lines
    ax.grid(b=True, which='minor', color='#008B8B', linestyle='-', alpha=0.2)

    for i in drehzahllist:
        ax.axvline(x = i, color = "gray") # Plotting a single vertical line
    ax.axvline(x = float(InfoValues[4]), color = "red") # Plotting a single vertical line

    ax.legend(loc='center left', bbox_to_anchor=(0, 1.12),ncol=10, fontsize = 8)

    fig0.tight_layout()
    fig0.show()
    ############################################################################
    fig1 = plt.figure(figsize=(26*cm, 13*cm))
    ax = fig1.add_subplot(111)
    ax.plot(speedList, powerList, label = 'Power_S1', color ='black')
    if 'S6' in graphen:
        ax.plot(speedList, powerList6, label = 'Power_S6 (40%, tS=2min)', linestyle = 'dashdot', color ='black')
    if 'Max-Werte' in graphen:
        ax.plot(speedList, powerListMax, label = 'Power_Max ', linestyle = 'dotted', color ='black')
    ax.set_xlabel('Speed [1/min]')
    ax.set_ylabel("Power [kW]")
    ax.set_xlim(xaxs_min,xaxs_max)
    ax.set_ylim(0,None)
    #ax.yaxis.set_major_locator(plticker.MultipleLocator(base=10))

    plt.suptitle(TitelInfo, fontsize=14, y =1.0)

    ax.grid(b=True, which='major', color='#666666', linestyle='-')#Gitternetz Show the major grid lines with dark grey lines
    ax.minorticks_on()# Show the minor grid lines with very faint and almost transparent grey lines
    ax.grid(b=True, which='minor', color='#008B8B', linestyle='-', alpha=0.2)

    for i in drehzahllist:
        ax.axvline(x = i, color = "gray") # Plotting a single vertical line
    ax.axvline(x = float(InfoValues[4]), color = "red") # Plotting a single vertical line

    ax.legend(loc='center left', bbox_to_anchor=(0, 1.12),ncol=10, fontsize = 8)

    fig1.tight_layout()
    fig1.show()
    ############################################################################
    fig2 = plt.figure(figsize=(26*cm, 13*cm))
    ax = fig2.add_subplot(111)
    ax.plot(speedList, torqueList, label = 'Torque_S1', color ='black')
    if 'S6' in graphen:
        ax.plot(speedList, torqueList6, label = 'Torque_S6 (40%, tS=2min)', linestyle = 'dashdot', color ='black')
    if 'Max-Werte' in graphen:
        ax.plot(speedList, torqueListMax, label = 'Torque_Max ', linestyle = 'dotted', color ='black')
    ax.set_xlabel('Speed [1/min]')
    ax.set_ylabel("Torque [Nm]")
    ax.set_xlim(xaxs_min,xaxs_max)
    ax.set_ylim(0,None)
    #ax.yaxis.set_major_locator(plticker.MultipleLocator(base=10))

    plt.suptitle(TitelInfo, fontsize=14, y =1.0)

    ax.grid(b=True, which='major', color='#666666', linestyle='-')#Gitternetz Show the major grid lines with dark grey lines
    ax.minorticks_on()# Show the minor grid lines with very faint and almost transparent grey lines
    ax.grid(b=True, which='minor', color='#008B8B', linestyle='-', alpha=0.2)

    for i in drehzahllist:
        ax.axvline(x = i, color = "gray") # Plotting a single vertical line
    ax.axvline(x = float(InfoValues[4]), color = "red") # Plotting a single vertical line

    ax.legend(loc='center left', bbox_to_anchor=(0, 1.12),ncol=10, fontsize = 8)

    fig2.tight_layout()
    fig2.show()

    ############################################################################
    fig3 = plt.figure(figsize=(26*cm, 13*cm))
    ax = fig3.add_subplot(111)
    ax.plot(speedList, currentList, label = 'Current_S1', color = '#BF2799')
    if 'S6' in graphen:
        ax.plot(speedList, currentList6, label = 'Current_S6 (40%, tS=2min)', linestyle = 'dashdot', color = '#BF2799')
    if 'Max-Werte' in graphen:
        ax.plot(speedList, currentListMax, label = 'Current_Max ', linestyle = 'dotted', color ='#BF2799')
    ax.set_xlabel('Speed [1/min]')
    ax.set_ylabel("Current [A]")
    ax.set_xlim(xaxs_min,xaxs_max)
    ax.set_ylim(0,None)
    #ax.yaxis.set_major_locator(plticker.MultipleLocator(base=10))

    ax2 = ax.twinx()
    ax2.plot(speedList, voltageList, label = 'Voltage_S1', color = '#16B2AD')
    if 'S6' in graphen:
        ax2.plot(speedList, voltageList6, label = 'Voltage_S6 (40%, tS=2min)', linestyle = 'dashdot', color ='#16B2AD')
    #if 'Max-Werte' in graphen:
    #   ax2.plot(speedList, voltageListMax, label = 'Voltage_Max', linestyle = 'dotted', color ='#16B2AD')
    ax2.set_ylabel("Voltage [V]")
    ax2.set_ylim(0,None)
    plt.suptitle(TitelInfo, fontsize=14, y =1.0)

    ax.grid(b=True, which='major', color='#666666', linestyle='-')#Gitternetz Show the major grid lines with dark grey lines
    ax.minorticks_on()# Show the minor grid lines with very faint and almost transparent grey lines
    ax.grid(b=True, which='minor', color='#008B8B', linestyle='-', alpha=0.2)

    for i in drehzahllist:
        ax.axvline(x = i, color = "green") # Plotting a single vertical line
    ax.axvline(x = float(InfoValues[4]), color = "red") # Plotting a single vertical line

    ax.legend(loc='center left', bbox_to_anchor=(0, 1.12),ncol=2, fontsize = 8)
    ax2.legend(loc='center right', bbox_to_anchor=(1.0, 1.12),ncol=2, fontsize = 8)
    #ax2.legend(loc='center left', bbox_to_anchor=(0.75, 1.12),ncol=1, fontsize = 8)

    fig3.tight_layout()
    fig3.show()

    ############################################################################
    fig4 = plt.figure(figsize=(26*cm, 13*cm))
    ax = fig4.add_subplot(111)
    ax.plot(speedList, powerList, label = 'Power_S1', color = '#C95200')
    if 'S6' in graphen:
        ax.plot(speedList, powerList6, label = 'Power_S6 (40%, tS=2min)', linestyle = 'dashdot', color ='#C95200')
    if 'Max-Werte' in graphen:
        ax.plot(speedList, powerListMax, label = 'Power_Max ', linestyle = 'dotted', color ='#C95200')
    ax.set_xlabel('Speed [1/min]')
    ax.set_ylabel("Power [kW]")
    ax.set_xlim(xaxs_min,xaxs_max)
    ax.set_ylim(0,None)
    #ax.yaxis.set_major_locator(plticker.MultipleLocator(base=10))

    ax2 = ax.twinx()
    ax2.plot(speedList, torqueList, label = 'Torque_S1', color = '#005A9A')
    if 'S6' in graphen:
        ax2.plot(speedList, torqueList6, label = 'Torque_S6 (40%, tS=2min)', linestyle = 'dashdot', color = '#005A9A')
    if 'Max-Werte' in graphen:
        ax2.plot(speedList, torqueListMax, label = 'Torque_Max ', linestyle = 'dotted', color ='#005A9A')
    ax2.set_ylabel("Torque [Nm]")
    ax2.set_ylim(0,None)
    plt.suptitle(TitelInfo, fontsize=14, y =1.0)

    ax.grid(b=True, which='major', color='#666666', linestyle='-')#Gitternetz Show the major grid lines with dark grey lines
    ax.minorticks_on()# Show the minor grid lines with very faint and almost transparent grey lines
    ax.grid(b=True, which='minor', color='#008B8B', linestyle='-', alpha=0.2)

    for i in drehzahllist:
        ax.axvline(x = i, color = "green") # Plotting a single vertical line
    ax.axvline(x = float(InfoValues[4]), color = "red") # Plotting a single vertical line
    #plt.axvline(x = 5, color = "green", label = "Index 5") # Plotting a single vertical line

    ax.legend(loc='center left', bbox_to_anchor=(0, 1.12),ncol=2, fontsize = 8)
    ax2.legend(loc='center right', bbox_to_anchor=(1.0, 1.12),ncol=2, fontsize = 8)
    #ax2.legend(loc='center left', bbox_to_anchor=(0.75, 1.12),ncol=1, fontsize = 8)

    fig4.tight_layout()
    fig4.show()

    return (fig0, fig1, fig2, fig3, fig4)

def CreateWord(variS1, variS6, plot4, plot5, InfoValues, savepath):
    Word_Path = eg.fileopenbox(msg = "Bitte die Diagramm-Vorlage auswählen", default = "*.docx")
    if Word_Path == None:
        sys.exit()
    docx_tpl = DocxTemplate(Word_Path)

    ZN = InfoValues[0]
    AS = InfoValues[1]
    DBL = InfoValues[2]
    MDB = InfoValues[3]
    Nenn = InfoValues[4]
    VD = InfoValues[5]
    IDVD = InfoValues[6]
    
    img1 = InlineImage(docx_tpl, plot4, height = Mm(77))
    img2 = InlineImage(docx_tpl, plot5, height = Mm(77))

    context = {'ZN' : ZN,
            'AS' : AS,
            'DBL' : DBL,
            'MDB' : MDB,
            'Nn' : Nenn,
            'VD' : VD,
            'IDVD' : IDVD,
            
            'now' : datetime.datetime.now(),
            'bild1' : img1,
            'bild2' : img2,
    }
    for row, col in itertools.product(variS1.index, variS1.columns):
        context[f'{row}_{col}'] = variS1.loc[row, col]

    for row, col in itertools.product(variS6.index, variS6.columns):
        context[f'{row}_{col}'] = variS6.loc[row, col]

    docx_tpl.render(context)

    default_filename = ('AS' + AS)

    # Save File
    saveyesno = save_file_dialogs(default_filename, savepath, extension = "docx")
    if saveyesno != None:
        while 1:    #überprüfen ob Worddate bereits geöffnet ist, ansonsten abspeichern
            try:
                docx_tpl.save(saveyesno)
                break
            except PermissionError:
                fileIsOpen = eg.buttonbox(msg="Die Word-Datei ist bereits geöffnet. Bitte schließen", choices = ("Abbrechen", "Ok") )
                if fileIsOpen == "Abbrechen":
                    eg.msgbox(msg='Abbruch durch Benutzer!', title='Benutzerabbruch', ok_button='OK', image=None, root=None)
                    sys.exit()
                else:
                    continue
        if eg.buttonbox(msg = "Protokolldatei wurde unter %s gespeichert.\n\n Soll die Datei geöffnet werden?" %(saveyesno),choices=('Ja', 'Nein')) == 'Ja':
            os.startfile(saveyesno)

    #docx_tpl.save('D:\VisualStudioProjekte\Dataframe für Diagramme\TestTemplate2.docx')

def main():
    faulheit = eg.choicebox(msg='zum Testen eine Vorbelegung auswählen oder Werte eingeben.', title='Pick an item', choices=['DC-0070038500 beschränkt auf 33,5kW', 'original Daten DC-0070038500', 'DC-0070119300', 'AC-0070074200', 'Werte eingeben'], preselect=0) #Faulheit
    #print(spinType) #Faulheit #debug

    graphen = eg.multchoicebox(msg='Welche Graphen sollen zusätzlich zu dem S1 Graph erzeugt werden?', title='Auswahl Graphen', choices=['S6', 'Max-Werte'], preselect=0, callback=None, run=True)
    if graphen == None:
        graphen = ""
    
    #Eingabe von allgemeinen Spindeldaten
    msgInfo = "Geben sie vorhandene Daten ein"
    titleInfo = "Spindel Info"
    fieldNamesInfo = ["Zeichnungsnummer", 
                    "AS-Nummer", 
                    "Motor_Hersteller_DBL", 
                    "Cytec-MDB", 
                    "Nenndrehzahl",
                    "Vorschaltdrossel mH",
                    "ID -Vorschaltdrossel"]
    if faulheit == 'DC-0070038500 beschränkt auf 33,5kW':
        fieldValues = ["CS/HSKA063/020W/200F-0993","0815","0038500","204-0815","2000","0,6","123-0815"]  # Vorbelegung der Eingabefelder # soll verwendet werden
    elif faulheit == 'original Daten DC-0070038500':
        fieldValues = ["CS/HSKA063/020W/200F-1234","4711","0038500orig","204-4711","2000","0,6","123-4711"] 
    elif faulheit == 'DC-0070119300':
        fieldValues = ["CS/HSKA063/020W/20000F-9999","x1x1","0070119300","204-x2x2","5700","0,7","123-x3x3"]
    elif faulheit == 'AC-0070074200':
        fieldValues = ["CS/HSKA063/999W/200F-1111","x4x4","0070074200","204-x5x5","1460","--","--"] 
    else:
        fieldValues = ["","","","","","",""]  # Vorbelegung der Eingabefelder # soll verwendet werden
    InfoValues = eg.multenterbox(msgInfo,titleInfo,fieldNamesInfo,fieldValues)


    ########### Wertevorgaben um nicht immer alles einzutippen später komplett auskommentieren  #################
    if faulheit == 'DC-0070038500 beschränkt auf 33,5kW' :
            # Daten 0070038500 beschränkt auf 33,5kW siehe ES-Datenblatt
        dafr1 = {0: {'Drehzahl_S1': '0', 'Spannung': '0', 'Strom': '125', 'Leistung': '0', 'Drehmoment': '160'}, 1: {'Drehzahl_S1': '2000', 'Spannung': '228', 'Strom': '125', 'Leistung': '33.5', 'Drehmoment': '160'}, 2: {'Drehzahl_S1': '4000', 'Spannung': '400', 'Strom': '68', 'Leistung': '33.5', 'Drehmoment': '80'}, 3: {'Drehzahl_S1': '8000', 'Spannung': '400', 'Strom': '54', 'Leistung': '33.5', 'Drehmoment': '40'}, 4: {'Drehzahl_S1': '12000', 'Spannung': '400', 'Strom': '76', 'Leistung': '33.5', 'Drehmoment': '26.7'}}
        spinType = 'Synchron'
        variS1 = pd.DataFrame(dafr1)
    elif faulheit == 'original Daten DC-0070038500':
            # original Daten 0070038500
        dafr1 = {0: {'Drehzahl_S1': '0', 'Spannung': '0', 'Strom': '125', 'Leistung': '0', 'Drehmoment': '160'}, 1: {'Drehzahl_S1': '2000', 'Spannung': '228', 'Strom': '125', 'Leistung': '33.5', 'Drehmoment': '160'}, 2: {'Drehzahl_S1': '4000', 'Spannung': '400', 'Strom': '110', 'Leistung': '58', 'Drehmoment': '138.5'}, 3: {'Drehzahl_S1': '8000', 'Spannung': '400', 'Strom': '90', 'Leistung': '58', 'Drehmoment': '69.2'}, 4: {'Drehzahl_S1': '12000', 'Spannung': '400', 'Strom': '105', 'Leistung': '58', 'Drehmoment': '46.2'}}
        spinType = 'Synchron'
        variS1 = pd.DataFrame(dafr1)
    elif faulheit == 'DC-0070119300':
            # original Daten 0070119300
        dafr1 = {0: {'Drehzahl_S1': '0', 'Spannung': '0', 'Strom': '95', 'Leistung': '0', 'Drehmoment': '67'}, 1: {'Drehzahl_S1': '5700', 'Spannung': '388', 'Strom': '95', 'Leistung': '40', 'Drehmoment': '67'}, 2: {'Drehzahl_S1': '7000', 'Spannung': '423', 'Strom': '75', 'Leistung': '40', 'Drehmoment': '54.6'}, 3: {'Drehzahl_S1': '24000', 'Spannung': '425', 'Strom': '66', 'Leistung': '40', 'Drehmoment': '15.9'}}
        spinType = 'Synchron'
        variS1 = pd.DataFrame(dafr1)
    elif faulheit == 'AC-0070074200':
        # original Daten AC 0070074200
        dafr1 = {0: {'Drehzahl_S1': '0', 'Spannung': '0', 'Strom': '0', 'Leistung': '0', 'Drehmoment': '0'}, 1: {'Drehzahl_S1': '1480', 'Spannung': '320', 'Strom': '180', 'Leistung': '74', 'Drehmoment': '480'}, 2: {'Drehzahl_S1': '1780', 'Spannung': '380', 'Strom': '157', 'Leistung': '74', 'Drehmoment': '397'}, 3: {'Drehzahl_S1': '6500', 'Spannung': '380', 'Strom': '157', 'Leistung': '74', 'Drehmoment': '109'}, 4: {'Drehzahl_S1': '8170', 'Spannung': '380', 'Strom': '104', 'Leistung': '50', 'Drehmoment': '58.4'}, 5: {'Drehzahl_S1': '9900', 'Spannung': '380', 'Strom': '77', 'Leistung': '37', 'Drehmoment': '35.6'}}
        spinType = 'Asynchron'
        variS1 = pd.DataFrame(dafr1)
    elif faulheit == 'Werte eingeben':
                #    # S1-Werte eingeben
        variS1 = eingabeS1() # normale Eingabe von Werten. Hier auskommentiert um Vorgaben dafr1 zu nutzen um nicht immer alles einzutippen. Wird später wieder aktiviert und die Vorgabe wird gelöscht.
        spinType = variS1[1] #Synchron/Asynchron
        print(variS1)# debug
        variS1 = variS1[0]
        # dafr1 = variS1.to_dict() # Umwandlung df in dict , hier aktuell nicht benötigt.
    print(variS1) #debug
    ###############################################################################################################

   
    ########### Wertevorgaben um nicht immer alles einzutippen        später komplett auskommentieren  #################
    if faulheit == 'DC-0070038500 beschränkt auf 33,5kW' :
                # Daten 0070038500 beschränkt auf 33,5kW siehe ES-Datenblatt
        dafr6 = {0: {'Drehzahl_S6': '0', 'Spannung_S6': '0', 'Strom_S6': '170', 'Leistung_S6': '0', 'Drehmoment_S6': '200.6'}, 1: {'Drehzahl_S6': '2000', 'Spannung_S6': '250', 'Strom_S6': '170', 'Leistung_S6': '42', 'Drehmoment_S6': '200.5'}, 2: {'Drehzahl_S6': '4000', 'Spannung_S6': '400', 'Strom_S6': '85', 'Leistung_S6': '42', 'Drehmoment_S6': '100.3'}, 3: {'Drehzahl_S6': '8000', 'Spannung_S6': '400', 'Strom_S6': '65', 'Leistung_S6': '42', 'Drehmoment_S6': '50.1'}, 4: {'Drehzahl_S6': '12000', 'Spannung_S6': '400', 'Strom_S6': '83', 'Leistung_S6': '42', 'Drehmoment_S6': '33.4'}}
        variS6 = pd.DataFrame(dafr6)
    elif faulheit == 'original Daten DC-0070038500':
            # original Daten 0070038500
        dafr6 = {0: {'Drehzahl_S6': '0', 'Spannung_S6': '0', 'Strom_S6': '165', 'Leistung_S6': '0', 'Drehmoment_S6': '200.5'}, 1: {'Drehzahl_S6': '2000', 'Spannung_S6': '250', 'Strom_S6': '165', 'Leistung_S6': '42', 'Drehmoment_S6': '200.5'}, 2: {'Drehzahl_S6': '4000', 'Spannung_S6': '400', 'Strom_S6': '144', 'Leistung_S6': '73', 'Drehmoment_S6': '174.3'}, 3: {'Drehzahl_S6': '8000', 'Spannung_S6': '400', 'Strom_S6': '120', 'Leistung_S6': '73', 'Drehmoment_S6': '87.1'}, 4: {'Drehzahl_S6': '12000', 'Spannung_S6': '400', 'Strom_S6': '117', 'Leistung_S6': '73', 'Drehmoment_S6': '58.1'}}
        variS6 = pd.DataFrame(dafr6)
    elif faulheit == 'DC-0070119300':
            # original Daten 0070119300
        dafr6 = {0: {'Drehzahl_S6': '0', 'Spannung_S6': '0', 'Strom_S6': '120', 'Leistung_S6': '0', 'Drehmoment_S6': '84'}, 1: {'Drehzahl_S6': '5700', 'Spannung_S6': '415', 'Strom_S6': '120', 'Leistung_S6': '50', 'Drehmoment_S6': '84'}, 2: {'Drehzahl_S6': '7000', 'Spannung_S6': '425', 'Strom_S6': '97', 'Leistung_S6': '50', 'Drehmoment_S6': '68'}, 3: {'Drehzahl_S6': '24000', 'Spannung_S6': '425', 'Strom_S6': '75', 'Leistung_S6': '50', 'Drehmoment_S6': '20'}}
        variS6 = pd.DataFrame(dafr6)
    elif faulheit == 'AC-0070074200':
            # original Daten AC 0070074200
        dafr6 = {0: {'Drehzahl_S6': '0', 'Spannung_S6': '0', 'Strom_S6': '0', 'Leistung_S6': '0', 'Drehmoment_S6': '0'}, 1: {'Drehzahl_S6': '1470', 'Spannung_S6': '320', 'Strom_S6': '218', 'Leistung_S6': '94', 'Drehmoment_S6': '611'}, 2: {'Drehzahl_S6': '1780', 'Spannung_S6': '380', 'Strom_S6': '187', 'Leistung_S6': '94', 'Drehmoment_S6': '504'}, 3: {'Drehzahl_S6': '5580', 'Spannung_S6': '380', 'Strom_S6': '207', 'Leistung_S6': '94', 'Drehmoment_S6': '161'}, 4: {'Drehzahl_S6': '6500', 'Spannung_S6': '380', 'Strom_S6': '157', 'Leistung_S6': '74', 'Drehmoment_S6': '109'}, 5: {'Drehzahl_S6': '9900', 'Spannung_S6': '380', 'Strom_S6': '77', 'Leistung_S6': '37', 'Drehmoment_S6': '35.6'}}
        variS6 = pd.DataFrame(dafr6) # dict in df wandeln.
    elif faulheit == 'Werte eingeben':
            #     # Eingabe bzw. berechnung der S6-Werte
                # Abfrage ob S6-Werte eingegeben werden oder berechnet werden
        s6Eingabeyesno = eg.ynbox(msg='S6-Werte eingeben oder berechnen?\nWenn Berechnen gewählt wird, werden die Werte mit 1,3 x S1-Werte errechnet.\nSmax-Werte werden mit 1,4 x S6-Werte berechnet.\nSpannungen werden nicht umgerechnet', title='S6-Werte eingeben oder berechnen ', choices=('Eingeben', 'Berechnen'), image=None, default_choice='Eingeben', cancel_choice='Berechnen')
        if s6Eingabeyesno == True:
                # wenn eingegeben werden soll
            variS6 = eingabeS6(spinType) # normale Eingabe von Werten. Hier auskommentiert um Vorgaben dafr6 zu nutzen um nicht immer alles einzutippen. Wird später wieder aktiviert und die Vorgabe wird gelöscht.
            print(variS6) # debug
            #dafr6 = variS6.to_dict() # Umwandlung df in dict , hier aktuell nicht benötigt.
        else:
                # wenn berechnet werden soll
            print("---vor Transpose-----S1--------") #debug
            print(variS1) #debug
            print(variS1.info()) #debug
                # kopieren der S1-Werte in S6-df und transponieren des df zur Berechnung (in Pandas werden anscheinend immer die Spalten und nicht die Zeilen bearbeitet)
            variS6 = variS1.copy(deep=True)
            variS6 = variS6.transpose()
            print("---nach Transpose-------------") #debug
            print(variS6) #debug
            print("---------------------------") #debug
            print(variS6.info()) #debug Info über df-Elemente

                # da die Werte als Objekte vorliegen müssen sie zum berechnen in float gewandelt werden
            variS6['Strom'] = variS6['Strom'].astype(float, errors = 'raise')
            variS6['Leistung'] = variS6['Leistung'].astype(float, errors = 'raise')
            variS6['Drehmoment'] = variS6['Drehmoment'].astype(float, errors = 'raise')
                # mit "transform" können Werte in df ausgetauscht werden.
            variS6['Strom'] = variS6['Strom'].transform(mult1_3) # multiplikation der S1-Werte (hier bereits in der Variablen variS6) mit 1,3
            variS6['Leistung'] = variS6['Leistung'].transform(mult1_3) # multiplikation der S1-Werte (hier bereits in der Variablen variS6) mit 1,3
            variS6['Drehmoment'] = variS6['Drehmoment'].transform(mult1_3) # multiplikation der S1-Werte (hier bereits in der Variablen variS6) mit 1,3
                # umbenenn der Spalten in S6 Namen da bis hier noch die S1 Spaltenbezeichnungen verwendet werden
            variS6.columns = ['Drehzahl_S6','Spannung_S6', 'Strom_S6', 'Leistung_S6', 'Drehmoment_S6']

                # zurück transponieren S6 in ursprüngliche Form.
            print("---vor Transpose-----S6--------") #debug
            print(variS6) #debug
            variS6 = variS6.transpose()
            print("---nach Transpose-----S6--------") #debug
            print(variS6) #debug
            
    print(variS6) #debug
    ####################################################################################################################

        #hier Smax Werte berechnen mit 1,4 x S6    mult1_4(x)
    variSmax = variS6.copy(deep=True) #kopieren der S6 Werte in neues DataFrame
    variSmax = variSmax.transpose()
        # da die Werte als Objekte vorliegen müssen sie zum berechnen in float gewandelt werden
    variSmax['Strom_S6'] = variSmax['Strom_S6'].astype(float, errors = 'raise')
    variSmax['Leistung_S6'] = variSmax['Leistung_S6'].astype(float, errors = 'raise')
    variSmax['Drehmoment_S6'] = variSmax['Drehmoment_S6'].astype(float, errors = 'raise')
    variSmax['Strom_S6'] = variSmax['Strom_S6'].transform(mult1_4) # multiplikation der S6-Werte (hier bereits in der Variablen variSmax) mit 1,4
    variSmax['Leistung_S6'] = variSmax['Leistung_S6'].transform(mult1_4) # multiplikation der S6-Werte (hier bereits in der Variablen variSmax) mit 1,4
    variSmax['Drehmoment_S6'] = variSmax['Drehmoment_S6'].transform(mult1_4) # multiplikation der S6-Werte (hier bereits in der Variablen variSmax) mit 1,4

        # umbenenn der Spalten in Smax Namen da bis hier noch die S6 Spaltenbezeichnungen verwendet werden
    variSmax.columns = ['Drehzahl_Smax','Spannung_Smax', 'Strom_Smax', 'Leistung_Smax', 'Drehmoment_Smax']

        # zurück transponieren Smax in ursprüngliche Form.
    # print("---vor Transpose-----Smax--------") #debug
    # print(variSmax) #debug
    variSmax = variSmax.transpose()
    # print("---nach Transpose-----Smax--------") #debug
    # print(variSmax) #debug

    # print("S1") #debug
    # print(variS1) #debug
    # print("S6") #debug
    # print(variS6) #debug
    # print("Smax") #debug
    # print(variSmax) #debug
    # print(variSmax.info()) #debug Info über df-Elemente


    # max value in Drehzahl_S1
    #max_speed = (variS1['Drehzahl_S1'].max())
    #print("max. eingegebene Drehzahl:  " + str(max_speed)) # max. eingegebene Drehzahl ermitteln


        # einzelne Serien aus dataframes erstellen, hier erst mal für S1 Werte
    Drehzahl = pd.to_numeric(variS1.iloc[0])
    Spannung = variS1.iloc[1].astype(float)
    Strom = variS1.iloc[2].astype(float)
    Leistung = variS1.iloc[3].astype(float)
    Drehmoment = variS1.iloc[4].astype(float)

    #print(Drehzahl, Leistung) #debug

        #aus eingegebenen Werten jeweils eine Liste erstellen
    drehzahllist = Drehzahl.tolist() # Drehzahlliste der eingegebenen S1 Werte
    spannunglist = dict(zip(drehzahllist, Spannung.tolist()))
    stromlist = dict(zip(drehzahllist, Strom.tolist())) # dict aus Drehzahl und Stromwerte zum hineinkopieren der eingegebenen Werte in Interpolations-Werte Liste
    leistunglist = dict(zip(drehzahllist, Leistung.tolist()))
    drehmomentlist = dict(zip(drehzahllist, Drehmoment.tolist()))
    #                               S6                                          #
        # einzelne Serien aus dataframes erstellen, hier für S6 Werte
    Drehzahl6 = pd.to_numeric(variS6.iloc[0])
    Spannung6 = variS6.iloc[1].astype(float)
    Strom6 = variS6.iloc[2].astype(float)
    Leistung6 = variS6.iloc[3].astype(float)
    Drehmoment6 = variS6.iloc[4].astype(float)

    #print(Drehzahl6, Leistung6) #debug

        #aus eingegebenen Werten jeweils eine Liste erstellen
    drehzahllist6 = Drehzahl6.tolist() 
    spannunglist6 = dict(zip(drehzahllist6, Spannung6.tolist())) 
    stromlist6 = dict(zip(drehzahllist6, Strom6.tolist())) # dict aus Drehzahl und Stromwerte zum hineinkopieren der eingegebenen Werte in Interpolations-Werte Liste
    leistunglist6 = dict(zip(drehzahllist6, Leistung6.tolist()))
    drehmomentlist6 = dict(zip(drehzahllist6, Drehmoment6.tolist()))
    #                               Smax                                          #
        # einzelne Serien aus dataframes erstellen, hier für S6 Werte
    DrehzahlMax = pd.to_numeric(variSmax.iloc[0])
    SpannungMax = variSmax.iloc[1].astype(float)
    StromMax = variSmax.iloc[2].astype(float)
    LeistungMax = variSmax.iloc[3].astype(float)
    DrehmomentMax = variSmax.iloc[4].astype(float)

    #print(DrehzahlMax, LeistungMax) #debug

        #aus eingegebenen Werten jeweils eine Liste erstellen
    drehzahllistMax = DrehzahlMax.tolist() 
    spannunglistMax = dict(zip(drehzahllistMax, SpannungMax.tolist())) 
    stromlistMax = dict(zip(drehzahllistMax, StromMax.tolist())) # dict aus Drehzahl und Stromwerte zum hineinkopieren der eingegebenen Werte in Interpolations-Werte Liste
    leistunglistMax = dict(zip(drehzahllistMax, LeistungMax.tolist()))
    drehmomentlistMax = dict(zip(drehzahllistMax, DrehmomentMax.tolist()))
    ####################################################################################################################################################

        # aus Drehzahlliste S1 und Drehzahlliste S6 eine gemeinsame Liste erstellen
    for i in drehzahllist6:
        if i in drehzahllist:
            pass
        else:
            drehzahllist.append(i)
    drehzahllist.sort(reverse = False)

        #Soll-Drehzahlenliste für interpolation und Graph erstellen
    step = 100
    maxspeeddiagram = Drehzahl.max() + step
    speedList = []

    for i in range(0,maxspeeddiagram,step):
        speedList.append(i)
        #Eingegebene Drehzahlen an Soll-Drehzahlliste anhängen
    for i in drehzahllist:
        if i in speedList:
            pass
        else:
            speedList.append(i)
        # Drehzahlenliste nach Größe sortieren
    speedList.sort(reverse = False)
    #print(speedList) #debug

    ############################################### S1 Berechnungen  ############################################################################

        # Zwischenwerte berechnen S1
    voltageList = []
    currentList = [] # Stromliste für Graph. Wird gefüllt durch interpolation und hineinkopieren der eingegebenen Werte
    powerList = []
    torqueList = []

            # Leistung S1, interpolation fehlender Werte bzw. hineinkopieren der eingegebenen Werte.
    for i in speedList:
        if i in drehzahllist and i in leistunglist:
            powerList.append(leistunglist[i])
        else:
            interpolate_x= i
            y_interp = interp1d(Drehzahl, Leistung)
            #print(Drehzahl) #debug
            #print("Leistungswert bei rpm = {} ist".format(interpolate_x), y_interp(interpolate_x)) #debug
            powerList.append(round((float(format(y_interp(interpolate_x)))), 2))
    # print("\n\n") #debug
    # print(speedList) #debug

            # Spannung S1
    for i in speedList:
            if i in drehzahllist and i in spannunglist:
                voltageList.append(spannunglist[i])
            else:
                interpolate_x= i    
                y_interp = interp1d(Drehzahl, Spannung)
                #print("Spannungswert bei rpm = {} ist".format(interpolate_x), y_interp(interpolate_x)) #debug
                voltageList.append(float(format(y_interp(interpolate_x))))

            # Strom S1
    for i in speedList:
            if i in drehzahllist and i in stromlist:
                currentList.append(stromlist[i])
            else:
                interpolate_x= i    
                y_interp = interp1d(Drehzahl, Strom)
                #print("Stromwert bei rpm = {} ist".format(interpolate_x), y_interp(interpolate_x)) #debug
                currentList.append(float(format(y_interp(interpolate_x))))


            # Drehmoment S1   Moment = Leistung * 9550 / Drehzahl
    sl = speedList # umsetzen der Liste in neue Variable weil im Durchlauf aus Liste ein Einzelwert wird!
    pl = powerList # umsetzen der Liste in neue Variable weil im Durchlauf aus Liste ein Einzelwert wird!
    for sl, pl in zip(sl, pl):
        if sl == 0: # wg. divison durch 0
            torqueList.append(drehmomentlist[0])
            #print("kopieren   Drehmomentwert bei rpm = " + str(sl) + " beträgt " + str(drehmomentlist[0]) + " Nm") #debug
        elif sl in drehzahllist and sl in drehmomentlist:
            torqueList.append(drehmomentlist[sl])
            #print("kopieren 2  Drehmomentwert bei rpm = " + str(sl) + " beträgt " + str(drehmomentlist[sl]) + " Nm") #debug
        else:
            torque = round((pl*9550/sl), 2)
            #print("Drehmomentwert bei rpm = " + str(sl) + " beträgt " + str(torque) + " Nm") #debug
            torqueList.append(torque)
            #print(torque) #debug

    ############################################### S6 Berechnungen   ############################################################################

        # Zwischenwerte berechnen S6
    voltageList6 = []
    currentList6 = []
    powerList6 = []
    torqueList6 = []

            # Leistung S6
    for i in speedList:
        if i in drehzahllist6 and i in leistunglist6:
            powerList6.append(leistunglist6[i])
        else:
            interpolate_x= i
            y_interp = interp1d(Drehzahl6, Leistung6)
            #print(Drehzahl6) #debug
            #print("Leistungswert_S6 bei rpm = {} ist".format(interpolate_x), y_interp(interpolate_x)) #debug
            powerList6.append(float(format(y_interp(interpolate_x))))
    #print("\n\n") #debug
    #print(powerList6) #debug


            # Spannung S6
    for i in speedList:
            if i in drehzahllist6 and i in spannunglist6:
                voltageList6.append(spannunglist6[i])
            else:
                interpolate_x= i    
                y_interp = interp1d(Drehzahl6, Spannung6)
                #print("Spannungswert_S6 bei rpm = {} ist".format(interpolate_x), y_interp(interpolate_x)) #debug
                voltageList6.append(float(format(y_interp(interpolate_x))))

            # Strom S6
    for i in speedList:
            if i in drehzahllist6 and i in stromlist6:
                currentList6.append(stromlist6[i])
            else:
                interpolate_x= i    
                y_interp = interp1d(Drehzahl6, Strom6)
                #print("Stromwert_S6 bei rpm = {} ist".format(interpolate_x), y_interp(interpolate_x)) #debug
                currentList6.append(float(format(y_interp(interpolate_x))))

            # Drehmoment S6   Moment = Leistung * 9550 / Drehzahl
    sl6 = speedList # umsetzen der Liste in neue Variable weil im Durchlauf aus Liste ein Einzelwert wird!
    pl6 = powerList6 # umsetzen der Liste in neue Variable weil im Durchlauf aus Liste ein Einzelwert wird!
    for sl6, pl6 in zip(sl6, pl6):
        if sl6 == 0: # wg. divison durch 0
            torqueList6.append(drehmomentlist6[0])
            #print("kopieren   Drehmomentwert bei rpm = " + str(sl6) + " beträgt " + str(drehmomentlist6[0]) + " Nm")
        elif sl6 in drehzahllist6 and sl6 in drehmomentlist6:
            torqueList6.append(drehmomentlist6[sl6])
            #print("kopieren 2   Drehmomentwert bei rpm = " + str(sl6) + " beträgt " + str(drehmomentlist6[sl6]) + " Nm") #debug
        else:
            torque6 = (pl6*9550/sl6)
            #print("Drehmomentwert_S6 bei rpm = " + str(sl6) + " beträgt " + str(torque6) + " Nm") #debug
            torqueList6.append(torque6)
            #print(torque) #debug
            
    ############################################### Smax Berechnungen   ############################################################################

        # Zwischenwerte berechnen Smax
    voltageListMax = []
    currentListMax = []
    powerListMax = []
    torqueListMax = []

            # Leistung Smax
    for i in speedList:
        if i in drehzahllistMax and i in leistunglistMax:
            powerListMax.append(leistunglistMax[i])
        else:
            interpolate_x= i
            y_interp = interp1d(DrehzahlMax, LeistungMax)
            #print(DrehzahlMax) #debug
            #print("Leistungswert_Smax bei rpm = {} ist".format(interpolate_x), y_interp(interpolate_x)) #debug
            powerListMax.append(float(format(y_interp(interpolate_x))))
    #print("\n\n") #debug
    #print(powerListMax) #debug


            # Spannung Smax
    for i in speedList:
            if i in drehzahllistMax and i in spannunglistMax:
                voltageListMax.append(spannunglistMax[i])
            else:
                interpolate_x= i    
                y_interp = interp1d(DrehzahlMax, SpannungMax)
                #print("Spannungswert_Smax bei rpm = {} ist".format(interpolate_x), y_interp(interpolate_x)) #debug
                voltageListMax.append(float(format(y_interp(interpolate_x))))

            # Strom Smax
    for i in speedList:
            if i in drehzahllistMax and i in stromlistMax:
                currentListMax.append(stromlistMax[i])
            else:
                interpolate_x= i    
                y_interp = interp1d(DrehzahlMax, StromMax)
                #print("Stromwert_Smax bei rpm = {} ist".format(interpolate_x), y_interp(interpolate_x)) #debug
                currentListMax.append(float(format(y_interp(interpolate_x))))

            # Drehmoment Smax   Moment = Leistung * 9550 / Drehzahl
    slMax = speedList # umsetzen der Liste in neue Variable weil im Durchlauf aus Liste ein Einzelwert wird!
    plMax = powerListMax # umsetzen der Liste in neue Variable weil im Durchlauf aus Liste ein Einzelwert wird!
    for slMax, plMax in zip(slMax, plMax):
        if slMax == 0: # wg. divison durch 0
            torqueListMax.append(drehmomentlistMax[0])
            #print("kopieren   Drehmomentwert bei rpm = " + str(slMax) + " beträgt " + str(drehmomentlistMax[0]) + " Nm") #debug
        elif slMax in drehzahllistMax and slMax in drehmomentlistMax:
            torqueListMax.append(drehmomentlistMax[slMax])
            #print("kopieren 2   Drehmomentwert bei rpm = " + str(slMax) + " beträgt " + str(drehmomentlistMax[slMax]) + " Nm") #debug
        else:
            torqueMax = (plMax*9550/slMax)
            #print("Drehmomentwert_Smax bei rpm = " + str(slMax) + " beträgt " + str(torqueMax) + " Nm") #debug
            torqueListMax.append(torqueMax)
            #print(torque) #debug

    print("S1") #debug
    print(variS1) #debug
    print("S6") #debug
    print(variS6) #debug
    print("Smax") #debug
    print(variSmax) #debug
    
        #     # überprüfen der MAX-Werte der Reihen
    # print('currentList')#debug
    # print(max(currentList))#debug
    # print('currentList6')#debug
    # print(max(currentList6))#debug
    # print('currentListMax')#debug
    # print(max(currentListMax))#debug
    
    # print('powerList')#debug
    # print(max(powerList))#debug
    # print('powerList6')#debug
    # print(max(powerList6))#debug
    # print('powerListMax')#debug
    # print(max(powerListMax))#debug
    
    # print('torqueList')#debug
    # print(max(torqueList))#debug
    # print('torqueList6')#debug
    # print(max(torqueList6))#debug
    # print('torqueListMax')#debug
    # print(max(torqueListMax))#debug

        #Daten lesen und Graph generieren
    #fig0, fig1, fig2, fig3, fig4 = make_plot(
    diagramme = make_plot(
                    InfoValues, 
                    Drehzahl, 
                    drehzahllist, 
                    speedList, 
                    currentList, 
                    currentList6, 
                    currentListMax, 
                    powerList, 
                    powerList6, 
                    powerListMax, 
                    voltageList, 
                    voltageList6, 
                    voltageListMax, 
                    torqueList,
                    torqueList6,
                    torqueListMax,
                    spinType,
                    graphen
                    )

        #Plots speichern
    savepath = eg.diropenbox(msg = "Speicherpfad für Diagramme auswählen\n --> Programm abbrechen mit 'Abbrechen'", title="Speicherpfad für Diagramme", default=None)
    if savepath == None:
        eg.msgbox(msg='Abbruch durch Benutzer!', title='Benutzerabbruch', ok_button='OK', image=None, root=None)
        sys.exit()

    plotname = ('current', 'power', 'torque', 'voltage-current', 'power-torque')
    i = 1
    for diagramme, plotname in zip(diagramme, plotname):
        default_filename = (InfoValues[1] + '-' + plotname)
        saveyesno = save_file_dialogs(default_filename, savepath, extension = "png")
        if saveyesno != None:
                if i == 4:
                    plot4 = saveyesno # Dateiname von voltage-current plot zu variable plot4 zu weisen für docxtpl
                elif i == 5:
                    plot5 = saveyesno
                diagramme.savefig(saveyesno)
                i += 1 
                print('saveyesnoDiag1 = '+ saveyesno)
                #diagname1 = saveyesno #diagname1 auf Rückgabewert aus Abspeichernfunktion setzen.
                # if eg.buttonbox(msg = "Protokolldatei wurde unter %s gespeichert.\n\n Soll die Datei geöffnet werden?" %(saveyesno),choices=('Ja', 'Nein')) == 'Ja':
                #     os.startfile(saveyesno)

     # Create the Wordfile:
    CreateWord(variS1, variS6, plot4, plot5, InfoValues, savepath)

    input("Warten auf's Ende\nbitte Enter drücken")
    #sys.exit()

if __name__ == '__main__':
    main()