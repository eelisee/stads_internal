#"hier --. einfügen" deklariert offene pfade

# import pandas as pd
import pdfkit
import csv
import datetime


# df = pd.read_csv('newmembers.csv')

# Definiere eine Abkürzung für den langen Pfad
wkhtmltopdf_path = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
# Verwende die Abkürzung in der Konfiguration
config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)



with open('newmembers5.csv', 'r') as df: #hier mitgliederliste einfügen
    reader = csv.reader(df)
    for row in reader:
        print(row)


current_time = datetime.datetime.now()





# HTML-Vorlage
html_template = """
<head>
    <meta charset="utf-8">
    <title>Lastschriftmandat STADS e.V.</title>
</head>


<h1 style="color: #5e9ca0;"><span style="color: #000000;">SEPA-Lastschriftmandat</span>
</h1> 
<p>&nbsp;</p>
<h4 style="color: #2e6c80;"><span style="color: #000000;">Zahlungsempf&auml;nger 
(Gl&auml;ubiger):</span></h4>
<p style="color: #2e6c80;"><span style="color: #000000;">STADS e.V.</span><br />
<span style="color: #000000;">Schloss<br /></span><span style="color: #000000;">
68131 Mannheim</span></p>
<p style="color: #2e6c80;"><span style="color: #000000;">Vertretungsberechtigter 
Vorstand:  Luca Marohn, Simon Schumacher, Elise Wolf</span></p>
<p style="color: #2e6c80;"><span style="color: #000000;">Gl&auml;ubiger ID: </span>
<span style="color: #000000;">DE03ZZZ00002109083</span></p>
<p style="color: #2e6c80;">&nbsp;</p>
<h4 style="color: #2e6c80;">
<span style="color: #000000;">Kontoinhaber (zahlungspflichtiges Mitglied):</span></h4>
<p><span style="color: #000000;">{First Name} {Second Name}<br />
</span><span style="color: #000000;">{Street}<br /></span>
<span style="color: #000000;">{Postal Code} {Coty}</span></p>
<p><span style="color: #000000;">Kontoverbindung: {Iban}</span></p>
<p><span style="color: #000000;">Mitgliedsnummer: {Mitgliedsnumme}</span></p>
<p><span style="color: #000000;">Mandatsreferenznummer: {Mandatsreferenz}</span></p>
<p>&nbsp;</p>
<p>Ich erm&auml;chtige den oben genannten Zahlungsempf&auml;nger, Zahlungen von meinem 
Konto mittels Lastschrift einzuziehen. Zugleich weise ich mein Kreditinstitut an, die 
vom oben genannten Zahlungsempf&auml;nger auf mein Konto gezogenen Lastschriften 
einzul&ouml;sen.</p>
<p>Hinweis: Ich kann innerhalb von acht Wochen, beginnend mit dem Belastungsdatum, 
die Erstattung des belasteten Betrages verlangen. Es gelten dabei die mit meinem 
Kreditinstitut vereinbarten Bedingungen. Ich bin damit einverstanden, dass zur 
Erleichterung des Zahlungsverkehrs, die grunds&auml;tzlich 14-t&auml;gige Frist 
f&uuml;r die Information vor Einzug einer f&auml;lligen Zahlung auf einen Tag vor 
Belastung verk&uuml;rzt wird.</p>
<p>&nbsp;</p>
<table style="height: 18px; width: 100%; border-collapse: collapse;">
<tbody>
<tr style="height: 10px;">
<td style="width: 26.373%; height: 18px;">________________</td>
<td style="width: 20.4072%; height: 18px;">________________
</td>
<td style="width: 53.2197%; height: 18px;">______________________________________</td>
</tr>
<tr style="height: 18px;">
<td style="width: 26.373%; height: 18px;">Ort</td>
<td style="width: 20.4072%; height: 18px;">Datum</td>
<td style="width: 53.2197%; height: 18px;">Unterschrift des Kontoinhabers</td>
</tr>
</tbody>
</table>
<p>&nbsp;</p>
"""


# Output-Pfad anpassen
output_path = r"stads_intra\mitgliedsaufnahme\op_ls" #hier outputpfad einfügen

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
        pdf_output_path = f'{output_path}\\lastschriftmandat_{row["Second Name"]}.pdf'
        pdfkit.from_string(filled_html, pdf_output_path, configuration=config)

        print(f"PDF erfolgreich generiert: {pdf_output_path}")