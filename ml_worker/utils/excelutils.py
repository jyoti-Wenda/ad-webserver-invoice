import os
import re
from datetime import datetime as dt
from dateutil.parser import parse
from thefuzz import fuzz

UPLOAD_FOLDER = '/flask_app/files/xlsx/'
import pandas as pd


def format_doc(doc_type, doc_name, extracted_data, pathfile):
    if doc_type == 'invoice':
        return format_invoice(doc_name, extracted_data)
    elif doc_type == 'loc':
        return loc(doc_name, extracted_data)
    else:
        return


def prune_text(text):
    chars = "\\`*_\{\}[]\(\)\|/<>#-\'\"+!$,\."
    for c in chars:
        if c in text:
            text = text.replace(c, "")
    return text


def cleanup_text(text):
    result = re.sub(r'[^a-zA-Z0-9]+', '', text)
    print('result',result)
    return result

def extract_gross_weight(text):
    result = re.sub(r'[^0-9.,]+', '', text)
    print('result', result)
    return result

'''
# Remove all non-numeric characters from the text
'''
def extract_numbers(text):
    result = re.sub(r'\D', '', text)
    print('Gross weight:', result)
    return result


def extract_numeric_values(text):
    result = re.findall(r'[-+]?\d*\.?\d+', text)
    print('result', result)
    return result

def extract_alphanumeric(text):
    result = re.findall(r'[a-zA-Z0-9]+', text)
    print('result', result)
    return result



def remove_leading_trailing_special_characters(text):
    result = re.sub(r'^[^a-zA-Z0-9]+|[^a-zA-Z0-9]+$', '', text)
    return result




def format_invoice(doc_name, extracted_data):
    """
    EXAMPLE extracted_data
    {
    "data_to_review": [
        [
        {
            "key": "Header",
            "page": 1,
            "type": "Inputs",
            "value": [
            {
                "key": "shipper",
                "state": "INCOMPLETE",
                "value": "Via  Irno,  221  I-84135  Salerno  -  Italy  -"
            },
            {
                "key": "consignee",
                "state": "INCOMPLETE",
                "value": "SOCIETE  ALGERIENNE  DE  PRODUCTION  DE  L  ELECTRICITE  SPE,  SPA  ROUTE  NATIONALE  N  38,  IMMEUBLE  DES  700  BEREAUX  GUE  DE  CONSTANTINE  KOUBA  ALGER  SAME  AS  CONSIGNEE"
            },
            {
                "key": "notify",
                "state": "INCOMPLETE",
                "value": ""
            },
            {
                "key": "incoterms",
                "state": "INCOMPLETE",
                "value": ""
            },
            {
                "key": "cad",
                "state": "INCOMPLETE",
                "value": ""
            },
            {
                "key": "container_type",
                "state": "INCOMPLETE",
                "value": ""
            },
            {
                "key": "container_id",
                "state": "INCOMPLETE",
                "value": ""
            },
            {
                "key": "seal_number",
                "state": "INCOMPLETE",
                "value": ""
            },
            {
                "key": "package_quantity",
                "state": "INCOMPLETE",
                "value": ""
            },
            {
                "key": "description",
                "state": "INCOMPLETE",
                "value": ""
            },
            {
                "key": "gross_weight",
                "state": "INCOMPLETE",
                "value": "Gross  Weight:  Kg.  710,00"
            },
            {
                "key": "hs_code",
                "state": "INCOMPLETE",
                "value": ""
            }
            ]
        }
        ]
    ],
    "detection_index": 0.8
    }
    """

    doc_name_contents = re.split("_", doc_name, 2)
    if len(doc_name_contents) == 3:
        attach_filename = doc_name_contents[2]
    else:
        attach_filename = doc_name
    reference_number = attach_filename.replace(".PDF", "").replace(".pdf", "")

    shipper, consignee, notify, incoterms, cad, container_type, container_id = [], [], [], [], [], [], []
    seal_number,package_quantity,description,gross_weight,hs_code=[],[],[],[],[]

    for page_nr in extracted_data:
        print(page_nr)
        data_to_review = extracted_data[page_nr]['data_to_review']
        for element in data_to_review:
            if element['key'] == 'Header':
                page = element['page']
                print('page: {}'.format(page))
                for element_item in element['value']:
                    # print(element_item['key'])
                    if element_item['key'] == 'shipper' and element_item['value'] != "":
                        value = remove_leading_trailing_special_characters(element_item['value'])
                        shipper.append(value)
                    if element_item['key'] == 'consignee' and element_item['value'] != "":
                        value = remove_leading_trailing_special_characters(element_item['value'])
                        consignee.append(value)
                    if element_item['key'] == 'notify' and element_item['value'] != "":
                        value = remove_leading_trailing_special_characters(element_item['value'])
                        notify.append(value)
                    if element_item['key'] == 'incoterms' and element_item['value'] != "":
                        value = remove_leading_trailing_special_characters(element_item['value'])
                        incoterms.append(value)
                    if element_item['key'] == 'cad' and element_item['value'] != "":
                        value = remove_leading_trailing_special_characters(element_item['value'])
                        cad.append(value)
                    if element_item['key'] == 'container_type' and element_item['value'] != "":
                        value = remove_leading_trailing_special_characters(element_item['value'])
                        container_type.append(value)
                    if element_item['key'] == 'container_id' and element_item['value'] != "":
                        value = remove_leading_trailing_special_characters(element_item['value'])
                        container_id.append(value)
                    if element_item['key'] == 'seal_number' and element_item['value'] != "":
                        value = extract_numbers(element_item['value'])
                        seal_number.append(value)
                    if element_item['key'] == 'package_quantity' and element_item['value'] != "":
                        value = extract_numbers(element_item['value'])
                        package_quantity.append(value)
                    if element_item['key'] == 'description' and element_item['value'] != "":
                        value = remove_leading_trailing_special_characters(element_item['value'])
                        description.append(value)
                    if element_item['key'] == 'gross_weight' and element_item['value'] != "":
                        value = extract_gross_weight(element_item['value'])
                        gross_weight.append(value)
                    if element_item['key'] == 'hs_code' and element_item['value'] != "":
                        value = extract_numeric_values(element_item['value'])
                        hs_code.append(value)

    xls_filepath = os.path.join(UPLOAD_FOLDER, reference_number + ".xlsx")
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    # CURRENT EXCEL HEADER
    # num. Booking | Tipo Cnt (Codice Iso) | Num. Container | Peso VGM | Nome persona autorizzata al Vgm | Caricatore/Shipper | Metodo 1 (Conservo scontrino) | Metodo 1 (allego scontrino) | Metodo 2(certificazione AEO o ISO9001/28000)
    df = pd.DataFrame({
                       'shipper': pd.Series(shipper),
                       'consignee': pd.Series(consignee),
                       'notify': pd.Series(notify),
                       'incoterms': pd.Series(incoterms),
                       'cad': pd.Series(cad),
                       'container_type': pd.Series(container_type),
                       'container_id': pd.Series(container_id),
                       'seal_number': pd.Series(seal_number),
                       'package_quantity': pd.Series(package_quantity),
                       'description': pd.Series(description),
                       'gross_weight': pd.Series(gross_weight),
                       'hs_code': pd.Series(hs_code)
                       })

    writer = pd.ExcelWriter(xls_filepath, engine='xlsxwriter')
    df.to_excel(writer, sheet_name="Sheet1", index=False)  # send df to writer
    worksheet = writer.sheets["Sheet1"]  # pull worksheet object
    for idx, col in enumerate(df):  # loop through all columns
        series = df[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width
    writer.close()

    # df.to_excel(xls_filepath, index=False)

    print('excel file has been created!')
    print(df)

    return xls_filepath, reference_number + ".xlsx"


def loc(doc_name, extracted_data):
    """
    EXAMPLE extracted_data
    {
  "data_to_review": [
    [
      {
        "key": "Header",
        "page": 1,
        "type": "Inputs",
        "value": [
          {
            "key": "lc_number",
            "state": "INCOMPLETE",
            "value": "0541ICD0000322099"
          },
          {
            "key": "date_of_issue",
            "state": "INCOMPLETE",
            "value": "220510"
          },
          {
            "key": "applicant",
            "state": "INCOMPLETE",
            "value": "BENBELLAT AHMED CITE LOMBARKIA 47 ROUTE DE TAZOULT 05000 BATNA ALGERIE"
          },
          {
            "key": "beneficiary",
            "state": "INCOMPLETE",
            "value": "BENETTI MACCHINE VIA PROVINCIALE NAZZANO 20 54033 CARRARA ITALY TEL 390585844347 FAX 390585842667"
          },
          {
            "key": "port_of_loading",
            "state": "INCOMPLETE",
            "value": "PORT ITALIEN"
          },
          {
            "key": "port_of_discharge",
            "state": "INCOMPLETE",
            "value": "PORT DE SKIKDA"
          },
          {
            "key": "description",
            "state": "INCOMPLETE",
            "value": "CFR PORT DE SKIKDA INCOTERMS 2020 01 HAVEUSE A CHAINE ET UN LOT DE PIECES DE RECHANGE MONTANT DE MARCHANDISE : EUR 130.000,00 MONTANT DU FRET : EUR 2.000,00 TOTAL : EUR 132.000,00 SUIVANT FACTURE PROFORMA NR 373 2021 Rev8 DU 04 05 2022"
          }
        ]
      }
    ]
  ],
  "detection_index": 0.8
}
"""

    doc_name_contents = re.split("_", doc_name, 2)
    if len(doc_name_contents) == 3:
        attach_filename = doc_name_contents[2]
    else:
        attach_filename = doc_name
    reference_number = attach_filename.replace(".PDF", "").replace(".pdf", "")

    lcNumber, dateOfIssue, Applicant, Beneficiary, portOfLoading = [], [], [], [], []
    portOfDischarge, latestDateOfShipment, Description = [], [], []
    for page_nr in extracted_data:
        print(page_nr)
        data_to_review = extracted_data[page_nr]['data_to_review']
        for element in data_to_review:
            if element['key'] == 'Header':
                page = element['page']
                print('page: {}'.format(page))
                for element_item in element['value']:
                    # print(element_item['key'])
                    if element_item['key'] == 'lc_number' and element_item['value'] != "":
                        # print('lc_number Before',lcNumber)
                        value = remove_leading_trailing_special_characters(element_item['value'])
                        lcNumber.append(value)
                    if element_item['key'] == 'date_of_issue' and element_item['value'] != "":
                        value = remove_leading_trailing_special_characters(element_item['value'])
                        dateOfIssue.append(value)
                    if element_item['key'] == 'applicant' and element_item['value'] != "":
                        value = remove_leading_trailing_special_characters(element_item['value'])
                        Applicant.append(value)
                    if element_item['key'] == 'beneficiary' and element_item['value'] != "":
                        value = remove_leading_trailing_special_characters(element_item['value'])
                        Beneficiary.append(value)
                    if element_item['key'] == 'port_of_loading' and element_item['value'] != "":
                        value = remove_leading_trailing_special_characters(element_item['value'])
                        portOfLoading.append(value)
                    if element_item['key'] == 'port_of_discharge' and element_item['value'] != "":
                        value = remove_leading_trailing_special_characters(element_item['value'])
                        portOfDischarge.append(value)
                    if element_item['key'] == 'latest_date_of_shipment' and element_item['value'] != "":
                        value = remove_leading_trailing_special_characters(element_item['value'])
                        latestDateOfShipment.append(value)
                    if element_item['key'] == 'description' and element_item['value'] != "":
                        value = remove_leading_trailing_special_characters(element_item['value'])
                        Description.append(value)

    xls_filepath = os.path.join(UPLOAD_FOLDER, reference_number + ".xlsx")
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    # CURRENT EXCEL HEADER
    # MERCE | TIPOLOGIA CONTAINER | SIGLA | COLLI | PESO LORDO | PESO NETTO | VOLUME | SIGILLI | TIPO | IMBALLO | TARA
    df = pd.DataFrame({
                       'lc_number': pd.Series(lcNumber),
                       'date_of_issue': pd.Series(dateOfIssue),
                       'applicant': pd.Series(Applicant),
                       'beneficiary': pd.Series(Beneficiary),
                       'port_of_loading': pd.Series(portOfLoading),
                       'port_of_discharge': pd.Series(portOfDischarge),
                       'latest_date_of_shipment': pd.Series(latestDateOfShipment),
                       'description': pd.Series(Description)
                       })

    writer = pd.ExcelWriter(xls_filepath, engine='xlsxwriter')
    df.to_excel(writer, sheet_name="Sheet1", index=False)  # send df to writer
    worksheet = writer.sheets["Sheet1"]  # pull worksheet object
    for idx, col in enumerate(df):  # loop through all columns
        series = df[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width
    writer.close()

    # df.to_excel(xls_filepath, index=False)

    print('excel file has been created!')
    print(df)

    return xls_filepath, reference_number + ".xlsx"