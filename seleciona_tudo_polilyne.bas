Attribute VB_Name = "SELE_ITEM_POLILYNE"
Sub nomedoobj()
    sSelecao = Format(Now(), "hh:mm:ss")
    Set ssetObj = ThisDrawing.SelectionSets.Add(sSelecao)
    ssetObj.SelectOnScreen
End Sub

Sub SELECIONA_ITENS()

Dim oEnt As AcadEntity
Dim Pt(0 To 2) As Double
Dim oLWP As AcadLWPolyline
Dim oP As AcadPolyline
Dim dblNewCords As Variant
Dim ssetObj As AcadSelectionSet
On Error Resume Next
ThisDrawing.SelectionSets.Item("TEST_SSET2").Delete


sSelecao = Format(Now(), "hh:mm:ss")
    Set ssetObj = ThisDrawing.SelectionSets.Add(sSelecao)
    ssetObj.SelectOnScreen



'ThisDrawing.Utility.GetEntity oEnt, Pt, "Select a polyline"
Set oLWP = ssetObj.Item(0) 'oEnt
dblCurCords = oLWP.Coordinates 'RECEBE TODAS AS COORDENADAS DA POLILYNE
iMaxCurArr = UBound(dblCurCords) 'RECEBE O NUMERO DE MATRIZES PRESENTE NA POLILYNE
iMaxNewArr = ((iMaxCurArr + 1) * 1.5) - 1 'AQUILE E FAZ O CALULO BATER COM O VALOR ESPERADO

ReDim dblNewCords(iMaxNewArr) As Double 'ATRIBUI A MATRIZ AO DBLNEWCORDS
iCurArrIdx = 0: iCnt = 1 'APENAS ATRIBUI VALORES A VARIAVES NÃO DEFINADAS
    For iNewArrIdx = 0 To iMaxNewArr 'INICIA O LOOP
    If iCnt = 3 Then
        dblNewCords(iNewArrIdx) = 0
        iCnt = 1
    Else
        dblNewCords(iNewArrIdx) = dblCurCords(iCurArrIdx)
        iCurArrIdx = iCurArrIdx + 1
        iCnt = iCnt + 1
    End If
    Next

Set ssetObj = ThisDrawing.SelectionSets.Add("TEST_SSET2")
ssetObj.SelectByPolygon acSelectionSetCrossingPolygon, dblNewCords


'FAZ ATRIBUIÇÃO AO BLOCO SELECIONADO

soma = 0
soma2 = 0
For j = 0 To ssetObj.Count - 1
    If ssetObj.Item(j).ObjectName = "AcDbText" Then
        
    soma = ssetObj.Item(j).TextString
    soma2 = soma2 + soma
    End If
Next

        For I = 0 To ssetObj.Count - 1
            If ssetObj.Item(I).ObjectName = "AcDbBlockReference" Then
                varAttributes = ssetObj.Item(I).GetAttributes 'recebe todos os atributos do bloco
                    For ii = 0 To UBound(varAttributes)
                            If UCase(varAttributes(ii).TagString) Like UCase("CTO-*") Then
                                
                                varAttributes(ii).TextString = "CTO-" & soma2
                            End If
                       
                    Next
            End If
        Next


End Sub
