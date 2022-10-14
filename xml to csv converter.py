#xml to csv converter
#TODO: add loop to find file names
#TODO: add loop to convert each file


import xml.etree.ElementTree as Xet
import pandas as pd
  
cols = ["name", "phone", "email", "date", "country"]
rows = []
fileNames = ["file1.xml", "file2.xml", "file3.xml"]

xmlparse = Xet.parse('filename goes here')
root = xmlparse.getroot()
for i in root:
    name = i.find("name").text
    phone = i.find("phone").text
    email = i.find("email").text
    date = i.find("date").text
    country = i.find("country").text
  
    rows.append({"name": name,
                 "phone": phone,
                 "email": email,
                 "date": date,
                 "country": country})
  
df = pd.DataFrame(rows, columns=cols)
  

df.to_csv('output.csv')
def main():
    pass

if __name__ == "__main__":
    main()