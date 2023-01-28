Attribute VB_Name = "poly_coordenadas"
Sub pl_coordenadas()

Dim ReturnObj As AcadObject
Dim BasePnt As Variant
Dim MyCoords As Variant


ThisDrawing.Utility.GetEntity ReturnObj, BasePnt, "Select an object"
If TypeOf ReturnObj Is AcadLWPolyline Then
    Set myobj = ReturnObj
    MyCoords = myobj.Coordinates
    mycoordscount = UBound(MyCoords)
    
    Dim ponto(0 To 8) As Double
    k = 0
    i = 0
    For x = 0 To mycoordscount - 1
        
        If i = 2 Then
        ponto(x) = 0
        k = k + 1
        i = 0
        End If
        ponto(k) = MyCoords(x)
    
        
        i = i + 1
        k = k + 1
        x = x + 1
    Next
Dim policopy As AcadPolyline
    Set policopy = ThisDrawing.ModelSpace.AddPolyline(ponto)
        
End If

End Sub
