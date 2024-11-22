import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os

# Load the Excel file
data = pd.read_excel('fdp2.xlsx')  # Updated file name

# Load the certificate template
template_path = '5G_new.png'  # Update with your actual path
output_folder = '5G FDP'  # Folder to save the certificates

# Font configuration
font_path = r'Poppins-Bold.ttf'  # Ensure this path is correct
font_size = 42  # Font size in pixels
try:
    font = ImageFont.truetype(font_path, font_size)
except OSError:
    print("Font not found. Using Arial as fallback.")
    font = ImageFont.truetype(r'C:\Windows\Fonts\arial.ttf', font_size)

# Text color
text_color = "red"

# Ensure the output folder exists
os.makedirs(output_folder, exist_ok=True)

# Iterate through the Excel data
for index, row in data.iterrows():
    # Concatenate Honor (unchanged) and Full Name (capitalized)
    full_name_with_honor = f"{row['HONOR']} {row['NAME'].upper()}"  # Only capitalize 'NAME'
    name_num=f"{row['NAME']} {row['NUM']}"

    # Load the certificate template
    template = Image.open(template_path)
    draw = ImageDraw.Draw(template)

    # Add Full Name with Honor
    name_position = (880, 715)  # Adjust this for placement
    draw.text(name_position, full_name_with_honor, fill=text_color, font=font)

    # Add College name (CLG) in uppercase
    college_name = row['CLG'].upper()  # Capitalize 'CLG'
    college_position = (340, 780)  # Position for College name
    draw.text(college_position, college_name, fill=text_color, font=font)

    # Save the certificate
    output_file = f"{output_folder}{name_num.replace(' ', '_')}.pdf"
    template.save(output_file)

print("Certificates generated successfully!")
