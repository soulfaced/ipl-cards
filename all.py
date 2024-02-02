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

certemplate = "all.pdf"
excelfile = "all.xlsx"
varname = "Name"
varmatches = "Matches"
varbaseprice = "Base Price"
varruns = "Runs"
varaverage = "Average"
varstrikerate = "Strike rate"
varnationality = "Nationality"
varrole = "Role"  # Add the 'Role' column
varimage = "Image"
varwickets = "Wickets"

# Provide default pixel dimensions for the text to be printed on the certificate
horz_name = 175
vert_name = 680

# Define default box dimensions
box_width = 200
box_height = 20

# Define default positions for each field
positions_dict = {
    varname: (horz_name, vert_name),
    varmatches: (590, 530 ),
    varbaseprice: (590, 690),
    varruns: (590, 410),
    varaverage: (590, 315),
    varstrikerate: (590, 100),
    varnationality: (200, 780),
    varrole: (175, 575),
    varwickets:(590,210),
}

# Define default position for the image
horz_image = 0
vert_image = 10

varfont = "Roboto-Bold.ttf"

# Specify different font sizes for each text element
fontsize_dict = {
    varname: 84,
    varmatches: 64,
    varbaseprice: 64,
    varruns: 64,
    varaverage: 64,
    varstrikerate: 64,
    varnationality: 32,
    varrole: 38,  
    varwickets:64,
}

# create the certificate directory
os.makedirs("all", exist_ok=True)

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
    role = row[varrole]  # Retrieve 'Role' from the row data
    Wickets = row[varwickets]

    # column contains the image URL
    image_url = row[varimage]

    # Download the image using requests
    image_data = requests.get(image_url).content

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(900,1000))

    # Draw the rectangular boxes for each field
    # for var, (horz, vert) in positions_dict.items():
    #     can.rect(horz, vert, box_width, box_height)

    # Set font and color for each field
    can.setFont("myFont", fontsize_dict[varname])  # Use the specified font size for 'Name'
    can.setFillColor(HexColor("#FAFBF9"))  # Set font color
    # Increase the vertical spacing between the two lines for the name
    line_spacing = 62

    # Split the name into words
    name_parts = name.split(maxsplit=1)

    # Draw the first word on the first line
    can.drawCentredString(horz_name + box_width / 2, vert_name - box_height / 2, name_parts[0])

    # Draw the rest of the name on the next line with increased spacing
    if len(name_parts) > 1:
        can.drawCentredString(horz_name + box_width / 2, vert_name - box_height / 2 - line_spacing, name_parts[1])



    # Adjust font size and provide text for other fields
    for var, fontsize in fontsize_dict.items():
        if var != varname:
            can.setFont("myFont", fontsize)
            can.drawCentredString(positions_dict[var][0] + box_width / 2, positions_dict[var][1] - box_height / 2,  str(row[var]))

    # Add image to the PDF
    if image_data:
        img = ImageReader(io.BytesIO(image_data))
        can.drawImage(img, horz_image, vert_image, width=580, height=580, mask='auto')  # Use 'mask' for PNG images

    can.save()
    packet.seek(0)
    new_pdf = PdfReader(packet)

    # Provide the certificate template
    existing_pdf = PdfReader(open(certemplate, "rb"))

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)
    destination = "all" + os.sep + name + ".pdf"
    outputStream = open(destination, "wb")
    output.write(outputStream)
    outputStream.close()
    print("created " + name + ".pdf")
