# openbom-parser

Steps to run

before exporting, **change values from fraction to decimal and remove any assemply types**

export excel sheet of parts from openbom only of Type Part
![alt text](./images/openbom-excel-file-example.png)


remove irrelavent rows and columns
![alt text](./images/cleaned-up-openbom-excel.png)

navigate to developer tab and click visual basic option
![alt text](./images/navigate-to-visual-basic.png)

click on sheet under VBAProject->Microsoft Excel Objects->Sheet
Then click on File->Import-> select heigh-length-swap.bas file from this project
a new modules folder should be created with the script imported
![alt text](./images/imported-script.png)


click on green button to run script, this should put all the smallest values in the height section
![alt text](./images/run-button-location.png)

export file as csv


# Steps to import into cutlist FX
Copy CSV formatted table into clip board
![alt text](./images/excel-copy.png)

open cutlist FX

under Edit drop down menu, click on `Import parts from clipboard`
![alt text](./images/cutlist-paste.png)

map the the csv headers to the cutlist fields, then hit finish
![alt text](./images/cutlist-mapping.png)

make sure to add material type once everything is imported!