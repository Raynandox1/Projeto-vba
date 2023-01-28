Attribute VB_Name = "minha_polyline"
Sub poli2()
Dim texto As acadText
Dim plineObj As AcadPolyline
Dim Points() As Double


Dim Start1 As Variant

Dim Finish1 As Variant

On Error GoTo tratar

Start1 = ThisDrawing.Utility.GetPoint(, "select")
Finish1 = ThisDrawing.Utility.GetPoint(Start1, "select point :")

k = 5
ReDim Points(k)
Points(0) = Start1(0): Points(1) = Start1(1): Points(2) = Start1(2)
Points(3) = Finish1(0): Points(4) = Finish1(1): Points(5) = Finish1(2)

Set plineObj = ThisDrawing.ModelSpace.AddPolyline(Points)


Dim meio(0 To 2) As Double


meio(0) = Finish1(0) - (Finish1(0) - Start1(0)) / 2

Set texto = ThisDrawing.ModelSpace.AddText(Round(plineObj.Length, 2), Finish1, 10)
    
    'desenha a polyline\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    While Start1(0) <> Finish1(0)
         Start1 = Finish1
         Finish1 = ThisDrawing.Utility.GetPoint(Start1, "select point :")
         plineObj.Delete
         
         
        ReDim Preserve Points(k + 3)
        Points(k + 1) = Finish1(0): Points(k + 2) = Finish1(1): Points(k + 3) = Finish1(2)
        Set plineObj = ThisDrawing.ModelSpace.AddPolyline(Points)
        
        k = k + 3
    Wend


    Exit Sub
tratar:
   Exit Sub
 End Sub
 
Sub poline()
    ' This example creates a polyline in model space.
On Error GoTo tratar
    Dim Start(0 To 2) As Double
    Dim Finish(0 To 2) As Double
    Dim Start1 As Variant
    Dim Finish1 As Variant
    Dim Dline As AcadLine
    Dim Dline2 As AcadLine
    
  Start(0) = 0

  
  Finish(0) = 1

  
    Start1 = ThisDrawing.Utility.GetPoint(, "select")
    Finish1 = ThisDrawing.Utility.GetPoint(Start1, "select point :")
    
    Set Dline = ThisDrawing.ModelSpace.AddLine(Start1, Finish1)
    
    

  'cria varias lines
    While Start(0) <> Finish(0)
         
         Start1 = Finish1
         Finish1 = ThisDrawing.Utility.GetPoint(Start1, "select point :")
         Set Dline2 = ThisDrawing.ModelSpace.AddLine(Start1, Finish1)
         Dline.Delete
        
        Start(0) = Start1(0)
        Finish(0) = Finish1(0)
    Wend
Exit Sub
tratar:
   Exit Sub
    
 End Sub
 
 
 
 
