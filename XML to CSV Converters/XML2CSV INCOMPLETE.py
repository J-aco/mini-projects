# Code to have a GUI that allows the user to select the columns to include in the CSV file.
# The code is incomplete and does not work.

import xml.etree.ElementTree as Xet
import csv
import os
from PySide6.QtWidgets import QMainWindow, QApplication, QWidget, QPushButton, QVBoxLayout, QLineEdit, QCheckBox, QLabel, QScrollArea

class XMLConverter(QWidget):
    def __init__(self):
        super().__init__()

        # Create a button to start the conversion process
        self.convert_button = QPushButton("Convert XML Files")
        self.convert_button.clicked.connect(self.convert_xml_files)

        # Create a line edit to enter the folder path
        self.folder_path_edit = QLineEdit()
        self.analyse_button = QPushButton("Analyse")
        self.analyse_button.clicked.connect(self.analyse_data)

        # Status Text
        self.status_text = QLabel("Waiting for folder path")
        self.sub_status_text = QLabel("")
        self.final_status_text = QLabel("")
        
        #analysis variables
        self.columns = []
        self.checkboxes=[]
        
        # Create a layout to hold the widgets
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.status_text)
        self.layout.addWidget(self.sub_status_text)
        self.layout.addWidget(self.folder_path_edit)
        self.layout.addWidget(self.analyse_button)
        self.layout.addWidget(self.final_status_text)

        # Set the layout of the window
        self.setLayout(self.layout)


    def analyse_data(self):
        # Parse one of the XML files to determine the columns that exist
        folder_path = self.folder_path_edit.text()
        file_names = os.listdir(folder_path)
        first_file = os.path.join(folder_path, file_names[0])
        #print(first_file)
        xmlparse = Xet.parse(first_file)
        root = xmlparse.getroot()

        # Create a list of the columns to include in the CSV file
        fields = [
            {
                'name': field.get('name'),
                'value': field.text
            }
        for field in root.findall('.//field')
        ]
        self.columns = [str(field['name']) for field in fields]
        
        # for column in self.columns:
        #     checkbox = QCheckBox(column)
        #     self.checkboxes.append(checkbox)
        #     self.layout.addWidget(checkbox)

        self.layout.addWidget(self.convert_button)
        self.status_text.setText("Select the columns to include in the CSV file")
        

    

    def convert_xml_files(self):
        # Get the folder path from the line edit
        folder_path = self.folder_path_edit.text()

        # Get the names of all the files in the folder
        file_names = os.listdir(folder_path)

        # Create a list of the selected columns
        selected_columns = self.columns
        # for i, checkbox in enumerate(self.checkboxes):
        #     if checkbox.isChecked():
        #         selected_columns.append(self.columns[i])

        with open('output.csv', 'w', newline='') as csvfile:
            # Create a DictWriter object to write the data to the CSV file
            writer = csv.DictWriter(csvfile, fieldnames=selected_columns)
            writer.writeheader()
            num_files = len(file_names)
            current_file = 0
            current_file_name = ""
            for file_name in file_names:
                current_file += 1
                current_file_name = file_name
                self.status_text.setText(f"Converting file {current_file} of {num_files}")
                # Check if the file is an XML file
                if file_name.endswith(".xml"):
                    # Parse the XML file and write the data to the CSV file
                    xmlparse = Xet.parse(os.path.join(folder_path, file_name))
                    root = xmlparse.getroot()
                    for i in root:
                        # Create a dictionary containing the XML data
                        data = {}
                        for column in selected_columns:
                            self.sub_status_text.setText(f"Current File: {current_file_name} | Processing column: {column}")
                            print(f"Current File: {current_file_name} | Processing column: {column}")
                            # Check if the element exists
                            element = i.find(column)
                            if element is not None:
                            # If the element exists, add its text to the data dictionary
                                data[column] = element.text
                            else:
                                # If the element doesn't exist, add an empty string to the data dictionary
                                data[column] = ""
                        # Write the dictionary to the CSV file
                        writer.writerow(data)
            self.status_text.setText(f"Converted {current_file} of {num_files}")
            self.sub_status_text.setText("")
            self.final_status_text.setText("Done!")
# Create an instance of the GUI and show it
app = QApplication([])
window = XMLConverter()
window.show()
app.exec()
