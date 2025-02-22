Attribute VB_Name = "Couleur_module"
Option Explicit
Option Base 1
Private module

Function HexToLongRGB(sHexVal As String) As Long() 'from stackoverflow
    Dim lRed As Long
    Dim lGreen As Long
    Dim lBlue As Long
    
    lRed = CLng("&H" & Left$(sHexVal, 2))
    lGreen = CLng("&H" & Mid$(sHexVal, 3, 2))
    lBlue = CLng("&H" & Right$(sHexVal, 2))
    
    Dim output(1 To 4) As Long
    output(1) = lRed
    output(2) = lGreen
    output(3) = lBlue
    output(4) = RGB(lRed, lGreen, lBlue)
    
    HexToLongRGB = output
End Function

Sub ApplyColor() 'regarde la feuille Main puis mettre la bonne couleur selon ID donné dans la colonne A
    WB_VBA_COLOR.Parent.ScreenUpdating = False
    WB_VBA_COLOR.Parent.Calculation = xlManual
    
    Dim ws As Worksheet
    Set ws = WS_MAIN
    
    Dim min As Integer
    Dim max As Integer
    Dim i As Integer
    
    min = 2
    
    max = ws.Range("A2501").End(xlUp).Row
    
    For i = min To max
        If ws.Cells(i, 1).Value <> "" And Len(ws.Cells(i, 1).Value) = 6 Then
            ws.Cells(i, 2).Interior.Color = HexToLongRGB(ws.Cells(i, 1).Value)(4)
            ws.Cells(i, 3).Value = HexToLongRGB(ws.Cells(i, 1).Value)(1)
            ws.Cells(i, 4).Value = HexToLongRGB(ws.Cells(i, 1).Value)(2)
            ws.Cells(i, 5).Value = HexToLongRGB(ws.Cells(i, 1).Value)(3)
        End If
    Next i
    
    WS_MAIN.Range("B1").Select
    WB_VBA_COLOR.Parent.ScreenUpdating = True
    WB_VBA_COLOR.Parent.Calculation = xlAutomatic
End Sub
