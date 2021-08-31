import csv
from lxml import etree
import os
import re
import shutil
import PySimpleGUI as sg

def surdegar(fil, katalog):
    # Läser in csv-filen
    fil = open(fil)
    csvfil = csv.reader(fil, dialect='excel', delimiter=';')

    # Skapar en lista för raderna i filen och lägger dessa till en lista
    rows = list()
    for row in csvfil:
        rows.append(row)

    # Tar bort rubrikraden
    del rows[0]

    # XML-del
    rotelement = etree.Element('LEVERANS')
    xmlfil = etree.ElementTree(rotelement)

    # Skapar variabler för att kunna jämföra med föregående rad, samt att säkerställa att fastighetsvärden inte dubbleras
    fastighetslista = list()
    fastighetslistamisc = list()
    lastrow = None
    rowcount = 1
    xmlrowcount = 0
    fellista = list()

    # TEST
    caselist = list()

    # Loopar igenom filen
    for row in rows:
        rowcount += 1
        if (row[0] == 'TOM') or (row[2] == 'TOM') or (row[3] == 'TOM'):
            print('VÄRDE SAKNAS', rowcount, row)
            fellista.append(f'VÄRDE SAKNAS {rowcount} {row}')
            continue
        if (row[0] == '') or (row[2] == '') or (row[3] == ''):
            print('VÄRDE SAKNAS RAD', rowcount, row)
            fellista.append(f'VÄRDE SAKNAS {rowcount} {row}')
            continue

        # Hoppar över tomma rader som saknar värde i fastighetskolumnen
        if row[0] == '':
            print(row)
            continue

        xmlrowcount += 1
        # Loop som genererar nytt ärende om inte ärendenummer och beslutsdatum är samma som föregående rad i csv-filen
        if lastrow != f'{row[2]}{row[3]}':
            # Rensar listor för att kunna jämföra fastigheter
            fastighetslista.clear()
            fastighetslistamisc.clear()
            # Uppdaterar jämförelsevariabeln
            lastrow = f'{row[2]}{row[3]}'

            # Kontrollerar om ärendet finns sedan tidigare i metadatafilen
            if lastrow in caselist:
                print('FEL! ÄRENDE FINNS REDAN!')
                print(rowcount, row)
                break

            # Hämtar värden att fylla elementen med
            fastighet = row[0].split(',')
            fastighetmisc = row[1].split(',')
            arendemening = row[4].upper()
            displayname = f'{row[2]}-{row[3]}, {arendemening}'
            beslutsdatum = row[2]
            beslutsnummer = row[3]
            arendestart = f'{beslutsdatum[:4]}-01-01'
            arendeslut = f'{beslutsdatum[:4]}-12-31'

            # Skapar xml på ärendenivå
            bygglov_element = etree.SubElement(rotelement, 'bygglov')
            idnumber = etree.SubElement(bygglov_element, 'ID').text = str(rowcount)
            displayname_element = etree.SubElement(bygglov_element, 'displayname').text = displayname
            arendemening_element = etree.SubElement(bygglov_element, 'arendemening').text = arendemening
            beslutsdatum_element = etree.SubElement(bygglov_element, 'beslutsdatum').text = beslutsdatum
            beslutsnummer_element = etree.SubElement(bygglov_element, 'beslutsnummer').text = beslutsnummer
            arendestart_element = etree.SubElement(bygglov_element, 'arendestart').text = arendestart
            arendeslut_element = etree.SubElement(bygglov_element, 'arendeslut').text = arendeslut
            fastigheter_element = etree.SubElement(bygglov_element, 'fastigheter')
            fastighetsmischead_element = etree.SubElement(bygglov_element, 'fastighetsinformation_ovrig')

            for values in fastighet:
                nyval = values.strip().upper()
                fastighet_element = etree.SubElement(fastigheter_element, 'fastighet').text = nyval
                fastighetslista.append(nyval.upper())

            for values in fastighetmisc:
                nyval = values.strip().upper()
                if nyval == '':
                    continue
                else:
                    fastighetmisc_element = etree.SubElement(fastighetsmischead_element, 'rad').text = nyval
                    fastighetslistamisc.append(nyval.upper())

            # Hämtar värden och skapar xml på handlingsnivå
            filnamn = row[6]
            displaynamehandling = f'{filnamn[:-4]}'
            handlingstyp = row[5].strip().replace('Ö', 'Övrigt').replace('R', 'Ritning')
            if handlingstyp == 'Övrigt':
                personuppgift = '10'
            elif handlingstyp == 'Ritning':
                personuppgift = '0'
            else:
                print(row, 'FEL VÄRDE PÅ HANDLINGSTYP')
                break

            handlingar_element = etree.SubElement(bygglov_element, 'handlingar')
            handling_element = etree.SubElement(handlingar_element, 'handling')
            displaynamehandling_element = etree.SubElement(handling_element, 'displayname').text = displaynamehandling
            filnamn_element = etree.SubElement(handling_element, 'filnamn').text = filnamn
            handlingstyp_element = etree.SubElement(handling_element, 'handlingstyp').text = handlingstyp
            personuppgiftsklassning_element = etree.SubElement(handling_element,'personuppgiftsklassning').text = personuppgift
            caselist.append(lastrow)

        # Loop om föregående rad tillhör samma ärende
        elif lastrow == f'{row[2]}{row[3]}':
            lastrow = f'{row[2]}{row[3]}'

            # Ärendenivå
            fastighet = row[0].split(',')
            fastighetmisc = row[1].split(',')

            for values in fastighet:
                nyval = values.strip().upper()
                if nyval in fastighetslista:
                    continue
                else:
                    fastighet_element = etree.SubElement(fastigheter_element, 'fastighet').text = nyval
                    fastighetslista.append(nyval)

            for values in fastighetmisc:
                nyval = values.strip().upper()
                if nyval in fastighetslistamisc:
                    continue
                elif nyval == '':
                    continue
                else:
                    fastighetmisc_element = etree.SubElement(fastighetsmischead_element, 'rad').text = nyval
                    fastighetslistamisc.append(nyval)

            # Handlingsnivå
            filnamn = row[6]
            displaynamehandling = f'{filnamn[:-4]}'
            handlingstyp = row[5].strip().replace('Ö', 'Övrigt').replace('R', 'Ritning')
            if handlingstyp == 'Övrigt':
                personuppgift = '10'
            elif handlingstyp == 'Ritning':
                personuppgift = '0'
            else:
                print(row, 'FEL VÄRDE PÅ HANDLINGSTYP')
                break
                

            handling_element = etree.SubElement(handlingar_element, 'handling')
            displaynamehandling_element = etree.SubElement(handling_element, 'displayname').text = displaynamehandling
            filnamn_element = etree.SubElement(handling_element, 'filnamn').text = filnamn
            handlingstyp_element = etree.SubElement(handling_element, 'handlingstyp').text = handlingstyp
            personuppgiftsklassning_element = etree.SubElement(handling_element,
                                                               'personuppgiftsklassning').text = personuppgift

    # Skapar xml-fil som lägger sig i den katalog som csv-filen finns.
    folderpath = katalog
    xmlfil.write(f'{folderpath}/output.xml', xml_declaration=True, encoding='utf-8', pretty_print=True)
    print(f'antal rader som genererade xml {xmlrowcount} i filen output.xml')

    if len(fellista) > 0:
        felfil = open(f'{folderpath}/fel.txt', 'w')
        felfil.write('\n'.join(fellista))

def kollafiler(fil, katalog):
    file = open(fil)
    csvfil = csv.reader(file, dialect='excel', delimiter=';')

    rows = list()
    fillista_csv = list()

    for row in csvfil:
        rows.append(row)

    for row in rows:
        if row[6] == '':
            continue

        filnamn = row[6]
        if filnamn in fillista_csv:
            print(f'{filnamn} dublett!')
            continue
        fillista_csv.append(filnamn)

    del fillista_csv[0]

    fillista_mapp = list()

    for file in os.scandir(katalog):
        if file.name.endswith('.tif'):
            fillista_mapp.append(file.name)

    result = Compare(fillista_mapp, fillista_csv)

    print(f' antal filer i csv: {len(fillista_csv)}')
    print(f' antal filer i mapp: {len(fillista_mapp)}')
    print(result)

def Compare(li1, li2):
    diff = [i for i in li1 + li2 if i not in li1 or i not in li2]
    return diff

def regexrattning(fil, katalog):
    file = open(fil)
    csvfil = csv.reader(file, dialect='excel', delimiter=';')
    folderpath = katalog

    # Skapar outputfilen corrected.csv som hamnar i samma katalog som originalfilen *.csv
    outfile = open(f'{folderpath}/corrected.csv', 'w')

    rows = []

    for row in csvfil:
        rows.append(row)

    for row in rows:
        match = re.match('(^\w{1,2}\s)[0]{1,3}(\d*)', row[3])
        if match:
            new = f'{match.group(1)}{match.group(2)}'
            correctedrow = f'{row[0]};{row[1]};{row[2]};{new};{row[4]};{row[5]};{row[6]};\n'
            print(correctedrow)
            outfile.write(correctedrow)
        else:
            oldrow = f'{row[0]};{row[1]};{row[2]};{row[3]};{row[4]};{row[5]};{row[6]};\n'
            outfile.write(oldrow)
            continue
    print('Färdig! Filen corrected.csv genererades.')
    

def flyttafilertillrest(fil, katalog):
    file = open(fil)
    csvfile = csv.reader(file, dialect='excel', delimiter=';')
    fillista = []
    rows = []

    for row in csvfile:
        rows.append(row)

    for row in rows:
        fillista.append(row[6])

    print(fillista)

    correct = 0
    incorrect = 0
    flyttadefiler = []
    flyttadefiler.clear()

    for root, dirs, files in os.walk(katalog):
        for name in files:
            sokvag = os.path.join(root, name)
            if (name in fillista) and (name not in flyttadefiler):
                try:
                    shutil.move(sokvag, r'J:\SBN surdegar\REST')
                    print('kopierar', name)
                    correct += 1
                    flyttadefiler.append(name)
                except shutil.SameFileError:
                    print('MISSLYCKAD!')
                    incorrect += 1
    print(f'Lyckade: {correct} \n Misslyckade: {incorrect}')


def kolladisplayname():
    pass

# GUI
def filvaljarfonster():
    layout2 = [[sg.Input(), sg.FileBrowse(key="-Fil-"), sg.Submit('Submit')]]
    window2 = sg.Window('Välj fil', layout2, size=(500, 100))
    while True:
        event, values = window2.read()
        if event == sg.WIN_CLOSED:
            break
        elif event == 'Submit':
            fil = values["-Fil-"]
            window2.close()
            return fil


sg.theme('LightGreen5')

layout = [[sg.Button("Fixa beslutsnummer"), sg.Button("Kontrollera filer"), sg.Button("Flytta filer"), sg.Button("Skapa arkivpaket")],[sg.Output(size=(125, 300), key=('_output_'), font='Consolas 10')]]


window = sg.Window('Surdegsfixer', layout, size=(700, 600), element_justification='c')

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    elif event == "Fixa beslutsnummer":
        fil = filvaljarfonster()
        if fil != None:
            try:
                window.FindElement('_output_').Update('')
                match = re.match('(.+)(\/.+?)$', fil)
                katalog = match.group(1)
                filtest = match.group(2)
                filextension = filtest[-3:]
                if filextension != 'csv':
                    print('Fel filformat! Välj en ".csv"')
                    continue 
                regexrattning(fil,katalog)
            except FileNotFoundError:
                print('Ingen fil vald')

    elif event == "Kontrollera filer":
        fil = filvaljarfonster()
        try:
            window.FindElement('_output_').Update('')
            match = re.match('(.+)(\/.+?)$', fil)
            katalog = match.group(1)
            filtest = match.group(2)
            filextension = filtest[-3:]
            if filextension != 'csv':
                print('Fel filformat! Välj en ".csv"')
                continue 
            kollafiler(fil, katalog)
        except AttributeError:
            print('Ingen fil vald!')

    elif event == "Skapa arkivpaket":
        fil = filvaljarfonster()
        if fil != None:
            try:
                window.FindElement('_output_').Update('')
                match = re.match('(.+)(\/.+?)$', fil)
                katalog = match.group(1)
                filtest = match.group(2)
                filextension = filtest[-3:]
                if filextension != 'csv':
                    print('Fel filformat! Välj en ".csv"')
                    continue 
                surdegar(fil,katalog)
            except FileNotFoundError:
                print('Ingen fil vald!')

    elif event == "Flytta filer":
        fil = filvaljarfonster()
        if fil != None:
            try:
                window.FindElement('_output_').Update('')
                match = re.match('(.+)(\/.+?)$', fil)
                katalog = match.group(1)
                filtest = match.group(2)
                filextension = filtest[-3:]
                if filextension != 'csv':
                    print('Fel filformat! Välj en ".csv"')
                    continue 
                flyttafilertillrest(fil, katalog)
            except AttributeError:
                print('Ingen fil vald!')
