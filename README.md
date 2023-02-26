# openbom-parser

Steps to run

before exporting, change values from fraction to decimal and remove any assemply types

export excel sheet of parts from openbom only of Type Part
![alt text](./images/openbom-excel-file-example.png)

change headers:
Name -> Label
Quantity -> Qty
Width (in) -> Width
Length (in) -> length
Height (in) -> Material

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

go to openbom and import csv file
