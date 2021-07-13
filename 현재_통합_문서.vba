Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call pausegame
End Sub

Private Sub Workbook_Open()
    Randomize
    Application.ScreenUpdating = False
    Cells(48, 48).Select
    Call ActiveSheet.Protect("Tkdlqjrj", False, True)
    Application.ScreenUpdating = True
End Sub
