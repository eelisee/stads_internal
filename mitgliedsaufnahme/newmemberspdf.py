#"hier --. einfügen" deklariert offene pfade

# import pandas as pd
import pdfkit
import csv

# df = pd.read_csv('newmembers5.csv')

# Definiere eine Abkürzung für den langen Pfad
wkhtmltopdf_path = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
# Verwende die Abkürzung in der Konfiguration
config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)

with open('newmembers5.csv', 'r') as df: #hier mitgliederliste einfügen
    reader = csv.reader(df)
    for row in reader:
        print(row)

# HTML-Vorlage
html_template = """
<head>
    <meta charset="utf-8">
    <title>Mitgliedsantrag STADS e.V.</title>
</head>

<h1>&nbsp;</h1>
<h1>Mitgliedsantrag STADS e.V.</h1>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>Hiermit beantrage ich meinen Beitritt als Mitglied zum Verein Students' Association 
for Data Analytics &amp; Statistics Mannheim (STADS) e.V.</p>
<p>&nbsp;</p>
<table border="0">
<tbody>
<tr>
<td style="width: 237.96875px;">Vorname</td>
<td style="width: 389.046875px;">{First Name}</td>
</tr>
<tr>
<td style="width: 237.96875px;">Nachname</td>
<td style="width: 389.046875px;">{Second Name}</td>
</tr>
<tr>
<td style="width: 237.96875px;">Stra&szlig;e, Hausnr.</td>
<td style="width: 389.046875px;">{Street}</td>
</tr>
<tr>
<td style="width: 237.96875px;">PLZ, Ort</td>
<td style="width: 389.046875px;">{Postal Code}, {Coty}</td>
</tr>
<tr>
<td style="width: 237.96875px;">Mailadresse</td>
<td style="width: 389.046875px;">{Email}</td>
</tr>
<tr>
<td style="width: 237.96875px;">Handynummer</td>
<td style="width: 389.046875px;">{Phone Number}</td>
</tr>
<tr>
<td style="width: 237.96875px;">Universit&auml;t / Hochschule</td>
<td style="width: 389.046875px;">{University Select}</td>
</tr>
<tr>
<td style="width: 237.96875px;">Fachrichtung</td>
<td style="width: 389.046875px;">{Subject Select}</td>
</tr>
<tr>
<td style="width: 237.96875px;">angestrebter Abschluss</td>
<td style="width: 389.046875px;">{Target Degree}</td>
</tr>
<tr>
<td style="width: 237.96875px;">Fachsemester (ohne Bachelorsemester)</td>
<td style="width: 389.046875px;">{Current Semester}</td>
</tr>
<tr>
<td style="width: 237.96875px;">Ressort</td>
<td style="width: 389.046875px;">{Ressort}</td>
</tr>
</tbody>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>Durch meine Unterschrift erkenne ich die Satzung und Vereinsordnung des Vereins an. 
Insbesondere stimme ich der internen Speicherung der hier angegebenen Daten zu.</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>______________________________________________________________</p>
<p>Ort, Datum, Unterschrift des Mitglieds</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p><sub>STADS e.V.<br /></sub><sub>Schloss<br />68131 Mannheim</sub></p>
<p><sub>Vertretungsberechtigter Vorstand: Luca Marohn, Simon Schumacher, Elise Wolf
</sub></p>
<p><sub>stads.uni-mannheim.de</sub></p>
"""


# Output-Pfad anpassen
output_path = r"\stads_intra\mitgliedsaufnahme\output" #hier outputpfad einfügen

# # Für jede Zeile in der Excel-Datei ein PDF erstellen
# for index, row in df.iterrows():
#     # Ersetze die Platzhalter im HTML-Template durch die entsprechenden Werte aus de
#     filled_html = html_template.format(**row)
    
#     # Generiere PDF und speichere es im angegebenen Ordner
#     pdf_output_path = f'{output_path}\\mitgliedsantrag_{df["Second Name"]}.pdf'
#     pdfkit.from_string(filled_html, pdf_output_path)

#     print(f"PDF erfolgreich generiert: {pdf_output_path}")


# Pfad zur CSV-Datei
csv_file_path = 'newmembers5.csv' #hier mitgliederliste einfügen

# Öffne die CSV-Datei mit dem Semikolon als Trennzeichen
with open(csv_file_path, 'r', newline='', encoding='ISO-8859-1') as file:
    reader = csv.DictReader(file, delimiter=';')
    
    # Für jede Zeile in der CSV-Datei
    for row in reader:
        # Ersetze die Platzhalter im HTML-Template durch die entsprechenden Werte CSV
        filled_html = html_template.format(**row)
        
        # Generiere PDF und speichere es im angegebenen Ordner
        pdf_output_path = f'{output_path}\\mitgliedsantrag_{row["Second Name"]}.pdf'
        pdfkit.from_string(filled_html, pdf_output_path, configuration=config)

        print(f"PDF erfolgreich generiert: {pdf_output_path}")