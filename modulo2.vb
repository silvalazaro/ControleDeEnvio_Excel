Sub capa1()
'
' capa1 Macro
'

Sheets("Capa1").Visible = True 'desoculta plan capa
Sheets("Capa1").Select 'seleciona capa
    
imprimir = Application.Dialogs(xlDialogPrinterSetup).Show 'seleciona impressora

If (imprimir) Then 'se imprimir for true

    ActiveWindow.SelectedSheets.PrintOut From:=1, To:=1, Copies:=1
    Application.ScreenUpdating = True 'imprimir
    
End If

Sheets("CAPA_EQUIP").Select 'volta pra plan CAPA_EQUIP

Sheets("Capa1").Visible = False 'oculta plan capa

End Sub

Sub capaEquip()
'
' capa1 Macro
'

Sheets("CapaEquip").Visible = True 'desoculta plan capa
Sheets("CapaEquip").Select 'seleciona capa
    
imprimir = Application.Dialogs(xlDialogPrinterSetup).Show 'seleciona impressora

If (imprimir) Then 'se imprimir for true

    ActiveWindow.SelectedSheets.PrintOut From:=1, To:=1, Copies:=1
    Application.ScreenUpdating = True 'imprimir
    
End If

    
Sheets("CAPA_EQUIP").Select 'volta pra plan CAPA_EQUIP

Sheets("CapaEquip").Visible = False 'oculta plan capa

End Sub
Sub capa2()
'
' capa2 Macro
'

Sheets("Capa2").Visible = True 'desoculta plan capa
Sheets("Capa2").Select 'seleciona capa
    
imprimir = Application.Dialogs(xlDialogPrinterSetup).Show 'seleciona impressora

If (imprimir) Then 'se imprimir for true

    ActiveWindow.SelectedSheets.PrintOut From:=1, To:=1, Copies:=1
    Application.ScreenUpdating = True 'imprimir
    
End If

Sheets("CAPA_EQUIP").Select 'volta pra plan CAPA_EQUIP
Sheets("Capa2").Visible = False 'oculta plan capa

End Sub

Sub capa3()
'
' capa3 Macro
'

Sheets("Capa3").Visible = True 'desoculta plan capa
Sheets("Capa3").Select 'seleciona capa
    
imprimir = Application.Dialogs(xlDialogPrinterSetup).Show 'seleciona impressora

If (imprimir) Then 'se imprimir for true

    ActiveWindow.SelectedSheets.PrintOut From:=1, To:=1, Copies:=1
    Application.ScreenUpdating = True 'imprimir
    
End If

    
Sheets("CAPA_EQUIP").Select 'volta pra plan CAPA_EQUIP

Sheets("Capa3").Visible = False 'oculta plan capa

End Sub

Sub capa4()
'
' capa4 Macro
'

Sheets("Capa4").Visible = True 'desoculta plan capa
Sheets("Capa4").Select 'seleciona capa
    
imprimir = Application.Dialogs(xlDialogPrinterSetup).Show 'seleciona impressora

If (imprimir) Then 'se imprimir for true

    ActiveWindow.SelectedSheets.PrintOut From:=1, To:=1, Copies:=1
    Application.ScreenUpdating = True 'imprimir
    
End If

Sheets("CAPA_EQUIP").Select 'volta pra plan CAPA_EQUIP

Sheets("Capa4").Visible = False 'oculta plan capa

End Sub
