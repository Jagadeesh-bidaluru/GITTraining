import os
from azure.core.credentials import AzureKeyCredential
from azure.ai.formrecognizer import DocumentAnalysisClient
from openpyxl import Workbook


# Read configuration from text file
with open("config.txt", "r") as config_file:
    config_data = config_file.readlines()

endpoint = config_data[0].strip()
key = config_data[1].strip()
Input_folder = config_data[2].strip()
output_folder = config_data[3].strip()

document_analysis_client = DocumentAnalysisClient(
    endpoint=endpoint, credential=AzureKeyCredential(key)
)

# Azure Form Recognizer endpoint and key
endpoint = "https://test-invoices-di.cognitiveservices.azure.com/"
key = "c35dcc1fa1c14a5cbec038897fe4ada3"
# Local file path
file_path = "/Users/L040495/PycharmProjects/InvoiceProcessing/inprogress/invoice.pdf"



document_analysis_client = DocumentAnalysisClient(
    endpoint=endpoint, credential=AzureKeyCredential(key)
)

# Analyze the document from file
with open(file_path, "rb") as f:
    # poller = document_analysis_client.begin_analyze_document("prebuilt-invoice", document=f)
    poller = document_analysis_client.begin_analyze_document("Invoice_Extraction", document=f)
    invoices = poller.result()

# Output Excel file path
output_excel_file = "/Users/L040495/PycharmProjects/InvoiceProcessing/output.xlsx"

# Check if the output file already exists
if os.path.exists(output_excel_file):
    # Load the existing workbook
    workbook = load_workbook(output_excel_file)
    sheet = workbook.active
else:
    # Create a new workbook if it doesn't exist
    workbook = Workbook()
    sheet = workbook.active
    # Write column headers if creating a new workbook
    headers = ["Vendor Name", "Customer Name", "Invoice ID", "Invoice Date", "Invoice Total", "PONumber"]
    sheet.append(headers)

# Process the extracted fields and append to Excel file
for idx, invoice in enumerate(invoices.documents):
    vendor_name = invoice.fields.get("VendorName").value if invoice.fields.get("VendorName") else ""
    customer_name = invoice.fields.get("CustomerName").value if invoice.fields.get("CustomerName") else ""
    invoice_id = invoice.fields.get("InvoiceId").value if invoice.fields.get("InvoiceId") else ""
    invoice_date = invoice.fields.get("InvoiceDate").value if invoice.fields.get("InvoiceDate") else ""
    invoice_total = invoice.fields.get("InvoiceTotal").value if invoice.fields.get("InvoiceTotal") else ""
    purchase_order = invoice.fields.get("PurchaseOrder").value if invoice.fields.get("PurchaseOrder") else ""

    row_data = [vendor_name, customer_name, invoice_id, invoice_date, invoice_total, purchase_order]
    sheet.append(row_data)

# Save the workbook
workbook.save(output_excel_file)

# Move the file to the completed folder
output_folder = "/Users/L040495/PycharmProjects/InvoiceProcessing/completed"
if not os.path.exists(output_folder):
    os.makedirs(output_folder)
os.rename(file_path, os.path.join(output_folder, os.path.basename(file_path)))

# Move the file to the completed folder
output_folder = "/Users/L040495/PycharmProjects/InvoiceProcessing/completed"
if not os.path.exists(output_folder):
    os.makedirs(output_folder)
os.rename(file_path, os.path.join(output_folder, os.path.basename(file_path)))


print("Output appended to", output_excel_file)

#re-write the code to read endpoint, key, file path,outputfolder from a text file named "config"

