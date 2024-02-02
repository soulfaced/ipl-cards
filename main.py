from PyPDF2 import PdfWriter, PdfReader
import io
import os
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor
from reportlab.lib.utils import ImageReader
import requests

certemplate = "temp.pdf"
excelfile = "data.xlsx"
varname = "Name"
varmatches = "Matches"
varbaseprice = "Base Price"
varruns = "Runs"
varaverage = "Average"
varstrikerate = "Strike rate"
varnationality = "Nationality"
varimage = "Image" 

# print("> Enter the pixel dimensions for the text to be printed on the certificate:")
horz_name = 25
vert_name = 650

# Define box dimensions here
box_width = 200
box_height = 20

horz_matches = 25
vert_matches = 600
horz_baseprice = 25
vert_baseprice = 550
horz_runs = 25
vert_runs = 500
horz_average = 25
vert_average = 450
horz_strikerate = 25
vert_strikerate = 400
horz_nationality = 25
vert_nationality = 350
horz_image = 300  # Adjust the horizontal position for the image
vert_image = 300  # Adjust the vertical position for the image

varfont = "Roboto-Bold.ttf"

# Specify different font sizes for each text element
fontsize_dict = {
    varname: 16,
    varmatches: 12,
    varbaseprice: 14,
    varruns: 12,
    varaverage: 14,
    varstrikerate: 12,
    varnationality: 14,
}

# create the certificate directory
os.makedirs("certificates", exist_ok=True)

# register the necessary font
pdfmetrics.registerFont(TTFont('myFont', varfont))

# provide the excel file that contains the participant information
data = pd.read_excel(excelfile)

for index, row in data.iterrows():
    name = row[varname].title().strip()
    matches = str(row[varmatches])
    base_price = str(row[varbaseprice])
    runs = str(row[varruns])
    average = str(row[varaverage])
    strike_rate = str(row[varstrikerate])
    nationality = row[varnationality]

    #  column contains the image URL
    image_url = row[varimage]

    # Download the image using requests
    image_data = requests.get(image_url).content

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)

    # Draw the rectangular boxes for each field
    can.rect(horz_name, vert_name, box_width, box_height)
    can.rect(horz_matches, vert_matches, box_width, box_height)
    can.rect(horz_baseprice, vert_baseprice, box_width, box_height)
    can.rect(horz_runs, vert_runs, box_width, box_height)
    can.rect(horz_average, vert_average, box_width, box_height)
    can.rect(horz_strikerate, vert_strikerate, box_width, box_height)
    can.rect(horz_nationality, vert_nationality, box_width, box_height)

    # Set font and color for each field
    can.setFont("myFont", fontsize_dict[varname])  # Use the specified font size for 'Name'
    can.setFillColor(HexColor("#FAFBF9"))  # Set font color

    # Provide the text for each field
    can.drawCentredString(horz_name + box_width / 2, vert_name + box_height / 2, "Name: " + name)

    # Adjust font size and provide text for other fields
    for var, fontsize in fontsize_dict.items():
        if var != varname:
            can.setFont("myFont", fontsize)
            # Define vert_dict here for other fields
            vert_dict = {
                varmatches: vert_matches,
                varbaseprice: vert_baseprice,
                varruns: vert_runs,
                varaverage: vert_average,
                varstrikerate: vert_strikerate,
                varnationality: vert_nationality,
            }
            can.drawCentredString(horz_name + box_width / 2, vert_dict[var] + box_height / 2, f"{var}: {row[var]}")

    # Add image to the PDF
    if image_data:
        img = ImageReader(io.BytesIO(image_data))
        can.drawImage(img, horz_image, vert_image, width=400, height=400, mask='auto')  # Use 'mask' for PNG images

    can.save()
    packet.seek(0)
    new_pdf = PdfReader(packet)

    # Provide the certificate template
    existing_pdf = PdfReader(open(certemplate, "rb"))

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)
    destination = "certificates" + os.sep + name + ".pdf"
    outputStream = open(destination, "wb")
    output.write(outputStream)
    outputStream.close()
    print("created " + name + ".pdf")
