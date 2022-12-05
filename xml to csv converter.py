import xml.etree.ElementTree as Xet
import csv
import os
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QLineEdit, QCheckBox

class XMLConverter(QWidget):
    def __init__(self):
        super().__init__()

        # Create a button to start the conversion process
        self.convert_button = QPushButton("Convert XML Files")
        self.convert_button.clicked.connect(self.convert_xml_files)

        # Create a line edit to enter the folder path
        self.folder_path_edit = QLineEdit()

        # Create a layout to hold the widgets
        layout = QVBoxLayout()
        layout.addWidget(self.folder_path_edit)
        layout.addWidget(self.convert_button)

        # Parse one of the XML files to determine the columns that exist
        folder_path = self.folder_path_edit.text()
        file_names = os.listdir(folder_path)
        xmlparse = Xet.parse(os.path.join(folder_path, file_names[0]))
        root = xmlparse.getroot()

        # Create a list of the columns to include in the CSV file
        self.columns = []
        for child in root:
            for subchild in child:
                if subchild.tag not in self.columns:
                    self.columns.append(subchild.tag)

        # Create a checkbox for each column
        self.checkboxes = []
        for column in self.columns:
            checkbox = QCheckBox(column)
            self.checkboxes.append(checkbox)
            layout.addWidget(checkbox)

        self.setLayout(layout)

    def convert_xml_files(self):
        # Get the folder path from the line edit
        folder_path = self.folder_path_edit.text()

        # Get the names of all the files in the folder
        file_names = os.listdir(folder_path)

        # Create a list of the selected columns
        selected_columns = []
        for i, checkbox in enumerate(self.checkboxes):
            if checkbox.isChecked():
                selected_columns.append(self.columns[i])

        with open('output.csv', 'w', newline='') as csvfile:
            # Create a DictWriter object to write the data to the CSV file
            writer = csv.DictWriter(csvfile, fieldnames=selected_columns)
            writer.writeheader()

            for file_name in file_names:
                # Check if the file is an XML file
                if file_name.endswith(".xml"):
                    # Parse the XML file and write the data to the CSV file
                    xmlparse = Xet.parse(os.path.join(folder_path, file_name))
                    root = xmlparse.getroot()
                    for i in root:
                        # Create a dictionary containing the XML data
                        data = {}
                        for column in selected_columns:
                            data[column] = i.find(column).text
                        # Write the dictionary to the CSV file
                        writer.writerow(data)

# Create an instance of the GUI and show it
app = QApplication([])
gui = XMLConverter()
gui.show()
app.exec_()