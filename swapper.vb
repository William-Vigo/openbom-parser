Sub sort()


'by Mal Rixon 21/09/2017
'ver 0.1
'Used to read output from "Create BOM" froum Fusion 360 to ensure material is arranged thickness, Length, Width as expected.
'Puts smallest value in Thickness largest value in Length Width should end up with the intermediate value
'Assume first row contains column headings.
'no blank rows

Dim sht As Worksheet
Dim LastRow As Long
Dim lastColumn As Integer
Dim rng As Range

Application.ReferenceStyle = xlR1C1
Set sht = ActiveSheet
'column headings must contain the following.  Change here if required.
WidthHeading = "Width"
LengthHeading = "Length"
HeightHeading = "Height"

 'find last row and column
'Using UsedRange
ActiveSheet.UsedRange 'Refresh UsedRange
  LastRow = sht.UsedRange.Rows.Count
  lastColumn = sht.UsedRange.Column - 1 + sht.UsedRange.Columns.Count
  
 'get column locations
Set rng = sht.Range(sht.Cells(1, 1), sht.Cells(1, lastColumn))

WidthCol = Application.Match(WidthHeading, rng, 0)
LengthCol = Application.Match(LengthHeading, rng, 0)
HeightCol = Application.Match(HeightHeading, rng, 0)

  'For each row in range

 'Set range from which to determine smallest value
For Row = 2 To LastRow
        Set rng = sht.Range(sht.Cells(Row, WidthCol), sht.Cells(Row, HeightCol))
        'Find minimum value which should be in height column
        'Worksheet function MIN returns the smallest value in a range
        'Worksheet function MATCH searches for a specified item in a range of cells, and then returns the relative position of that item in the range
         HeightREf = Application.Match(Application.Min(rng), rng, 0)

         
         
         
        'check if in height column
        If HeightREf <> 3 Then
                
           'If not swap for value in height column
           'first save  length value
           temp = sht.Cells(Row, HeightREf + WidthCol - 1)
           'second swap value curently in height column to where height data was
           sht.Cells(Row, HeightREf + WidthCol - 1) = sht.Cells(Row, HeightCol)
           'then save hegiht data to correct cell
           sht.Cells(Row, HeightCol) = temp
         End If
         
           
        'Find Maximum value and assign to length
        LengthREf = Application.Match(Application.Max(rng), rng, 0)
         
         ' Check if in Length column
         If LengthREf <> 2 Then
         
          'If not swap for value in height column
           'first save  length value
           temp = sht.Cells(Row, LengthREf + WidthCol - 1)
           'second swap value curently in height column to where height data was
           sht.Cells(Row, LengthREf + WidthCol - 1) = sht.Cells(Row, LengthCol)
           'then save hegiht data to correct cell
           sht.Cells(Row, LengthCol) = temp
           
        End If
Next Row

End Sub