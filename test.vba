```Visual Basic

Sub daydelete()
 
 Dim i As Long
 
 Dim sh As Worksheet
 
 
 Set sh = Worksheets("あさぎり01")
 
 
  
  
  For i = 1 To 31
  
  
  Set sh = Worksheets("あさぎり" & Format(i, "00"))
  
  
  
  sh.Range("B5:E6").Value = ""
  
  sh.Range("A11:H23").Value = ""
  
  sh.Range("A28:E34").Value = ""
  
  sh.Range("H26").Value = ""
  
  
  sh.Range("B38:I51").Value = ""
  
  sh.Range("B54:C54").Value = ""
  
  sh.Range("D55:E55").Value = ""
  
  
  
  Next i
  
  
 
 
 sh.Select
 
 End Sub

```
