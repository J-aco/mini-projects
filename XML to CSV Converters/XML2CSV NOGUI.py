# XML to CSV Converter
# Works for XML  files with the following structure:
# <root>
#     <field name="column1">value1</field>
#     <field name="column2">value2</field>
#     <field name="column3">value3</field>
# </root>
#
# Asks for the folder name containing the XML files to convert and 
# saves the CSV file at the same directory as the folder with the same name as the folder.
#
# Remember to change the directory paths to match your system.

import csv
import os
import xml.etree.ElementTree as ET


print("XML to CSV Converter")
print("Enter the folder name containing the XML files to convert")
folder = str(input('>>>'))

# Set the directory containing the xml files
xml_dir = f'C:\\Users\\User\\Path\\ToFile\\{folder}'

# Set the path to save the csv file
csv_file_path = f'C:\\Users\\User\\Path\\ToFile\\{folder}.csv'

# Open the csv file for writing. Line terminator is used to stop blank lines being written.
csv_file = open(csv_file_path, 'w', encoding='UTF8')
csv_writer = csv.writer(csv_file, lineterminator='\n')

# Set a flag to indicate if the column headers have been written
headers_written = False

# Loop through all xml files in the xml directory
for xml_file in os.listdir(xml_dir):
    # Only process xml files
    if xml_file.endswith(".xml"):
        print(f"Processing file: {xml_file}")
        tree = ET.parse(os.path.join(xml_dir, xml_file))
        root = tree.getroot()

        # If this is the first xml file and the headers have not been written, write the field names as the column headers in the csv file
        if not headers_written:
            csv_writer.writerow([field.attrib['name'] for field in root.findall('.//field')])
            headers_written = True
        
        csv_writer.writerow([field.text for field in root.findall('.//field')])

csv_file.close()