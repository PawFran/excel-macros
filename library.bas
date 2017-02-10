Attribute VB_Name = "library"
Sub autofit()
Attribute autofit.VB_ProcData.VB_Invoke_Func = "A\n14"
    ' ctrl+shift+a
    Cells.EntireColumn.autofit
End Sub

Sub pasteValues()
Attribute pasteValues.VB_ProcData.VB_Invoke_Func = "V\n14"
'
' pasteValues Makro
'
' Klawisz skrótu: Ctrl+Shift+V
'
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Application.CutCopyMode = False
End Sub


Sub deleteAllSelected()
'
' deleteAllSelected Makro
'
' Klawisz skrótu: Ctrl+Shift+D
'
    Selection.ClearContents

End Sub
Sub pasteFormats()
Attribute pasteFormats.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' pasteFormats Makro
'
' Klawisz skrótu: Ctrl+Shift+F
'
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    Application.CutCopyMode = False
End Sub
Sub pasteFormulas()
Attribute pasteFormulas.VB_ProcData.VB_Invoke_Func = "Z\n14"
'
' pasteFormulas Makro
'
' Klawisz skrótu: Ctrl+Shift+Z
'
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End Sub
Sub copyPasteValues()
Attribute copyPasteValues.VB_ProcData.VB_Invoke_Func = "C\n14"
'
' copyPasteValues Makro
'
' Klawisz skrótu: Ctrl+Shift+C
'
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
End Sub
Sub autofilterOnOff()
Attribute autofilterOnOff.VB_ProcData.VB_Invoke_Func = "Q\n14"
'
' autofilterOnOff Makro
'
' Klawisz skrótu: Ctrl+Shift+F
'
    Selection.AutoFilter
    
End Sub

Sub fillEmpty()
Attribute fillEmpty.VB_ProcData.VB_Invoke_Func = "E\n14"
    ' Ctrl + Shift + E
    ' fill from above
    For Each myCell In Selection
        If myCell.Value = "" Then
            myCell.Value = myCell.Offset(-1, 0).Value
        End If
    Next

End Sub

Sub fillEmptyFromBelow()
    ' Ctrl + E
    For Each myCell In Selection
        If myCell.Value = "" Then
            myCell.Value = myCell.Offset(1, 0).Value
        End If
    Next

End Sub

Sub removeDuplicates()
Attribute removeDuplicates.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' Keyboard Shortcut: Ctrl+Shift+R
'
    Selection.removeDuplicates Columns:=1, Header:=xlNo
End Sub

Sub test()
    
End Sub


