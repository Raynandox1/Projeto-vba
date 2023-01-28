Attribute VB_Name = "ESQUEMATICO_V2"
Sub ESQUEMATICO()

Dim sPath_Bloco As String
Dim Pinicial As Variant
Dim blockRef As AcadBlockReference
Dim coordenada(0 To 2) As Double
Dim texto As acadText

Set acad = GetObject(, "AutoCAD.Application")
If Err <> 0 Then
    Set acad = GetObject("AutoCAD.Application")
    acad.Visible = True
    MsgBox "abra o autocad e reinicie essa macro."
    Exit Sub
End If

Set doc = acad.ActiveDocument

' retorna o ponto usando o prompet
Pinicial = doc.Utility.GetPoint(, "selecione um ponto: ")

Dim ramais As Integer
Dim sp As Integer


ramais = InputBox("Quantos ramais")
sp = InputBox("Sp inicial")
pulo = 0


'bloco maior\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
sPath_Bloco = "C:\Users\rayna\OneDrive\햞ea de Trabalho\DRIVE NUVEM\projetos em vba\esquematico\bloco maior.dwg"

Set blockRef = doc.ModelSpace.InsertBlock(Pinicial, sPath_Bloco, 1, 1, 1, 0)
    blockRef.Explode
    blockRef.Delete
    
For I = 1 To ramais
    
'linha maior\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  sPath_Bloco = "C:\Users\rayna\OneDrive\햞ea de Trabalho\DRIVE NUVEM\projetos em vba\esquematico\linha maior.dwg"
    coordenada(0) = Pinicial(0) + 14.01
    coordenada(1) = Pinicial(1) - 11.25 - pulo
    coordenada(2) = Pinicial(2)
    
    Set blockRef = doc.ModelSpace.InsertBlock(coordenada, sPath_Bloco, 1, 1, 1, 0)
    blockRef.Explode
    blockRef.Delete
    
    'texto\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    

    
    coordenada(0) = Pinicial(0) + 14.06
    coordenada(1) = Pinicial(1) - 10.54 - pulo
    coordenada(2) = Pinicial(2)
    
    Set texto = ThisDrawing.ModelSpace.AddText("TR01-F" & sp, coordenada, 1.5)
    'XLnCADText.Rotation = (90 * 3.14) / 180
    texto.StyleName = "Times New Roman" 'ESTILO
    texto.ScaleFactor = 0.7 'ESCALA
    texto.Layer = "CTO_60-40"
    
    'texto cto e bloco quadrado e triangulo
    reservas = InputBox("Quantas reservas na sp-" & sp)
            salto = 0
            SUB1 = 0
            For II = 1 To 8 - reservas
            
            
            'texto cto\\\\\\\\\\\\\
                    coordenada(0) = Pinicial(0) + 26.96 + salto
                    coordenada(1) = Pinicial(1) - 10.94 - pulo
                    coordenada(2) = Pinicial(2)
                    
                    Set texto = ThisDrawing.ModelSpace.AddText("CTO-" & sp * 8 - SUB1, coordenada, 1.5)
                    texto.Rotation = (90 * 3.14) / 180
                    texto.StyleName = "Times New Roman" 'ESTILO
                    texto.ScaleFactor = 0.7 'ESCALA
                    texto.Layer = "CTO_60-40"
                        
            'quadrado menor\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    sPath_Bloco = "C:\Users\rayna\OneDrive\햞ea de Trabalho\DRIVE NUVEM\projetos em vba\esquematico\quadrado menor.dwg"
                      coordenada(0) = Pinicial(0) + 27.05 + salto
                      coordenada(1) = Pinicial(1) - 11.25 - pulo
                      coordenada(2) = Pinicial(2)
                      
                      Set blockRef = doc.ModelSpace.InsertBlock(coordenada, sPath_Bloco, 1, 1, 1, 0)
                      blockRef.Explode
                      blockRef.Delete
                             
            salto = salto + 3.34
            SUB1 = SUB1 + 1
            Next
            For III = 1 To reservas
                    
                    'texto cto\\\\\\\\\\\\\
                    coordenada(0) = Pinicial(0) + 26.96 + salto
                    coordenada(1) = Pinicial(1) - 10.94 - pulo
                    coordenada(2) = Pinicial(2)
                    
                    Set texto = ThisDrawing.ModelSpace.AddText("RES-" & sp * 8 - SUB1, coordenada, 1.5)
                    texto.Rotation = (90 * 3.14) / 180
                    texto.StyleName = "Times New Roman" 'ESTILO
                    texto.ScaleFactor = 0.7 'ESCALA
                    texto.Layer = "CTO_60-40"
                    texto.color = acMagenta
                    'quadrado menor\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    sPath_Bloco = "C:\Users\rayna\OneDrive\햞ea de Trabalho\DRIVE NUVEM\projetos em vba\esquematico\triangulo.dwg"
                      coordenada(0) = Pinicial(0) + 27.05 + salto
                      coordenada(1) = Pinicial(1) - 11.25 - pulo
                      coordenada(2) = Pinicial(2)
                      
                      Set blockRef = doc.ModelSpace.InsertBlock(coordenada, sPath_Bloco, 1, 1, 1, 0)
                      blockRef.Explode
                      blockRef.Delete
                   
            salto = salto + 3.34
            SUB1 = SUB1 + 1
            Next
    
    sp = sp + 1
    pulo = pulo + 14.18
Next
    





Set acad = Nothing
End Sub
