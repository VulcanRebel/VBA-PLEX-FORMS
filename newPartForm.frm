VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} newPartForm 
   Caption         =   "New Part Creation Form"
   ClientHeight    =   11640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7140
   OleObjectBlob   =   "newPartForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "newPartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label10_Click()

End Sub

Private Sub newPartClear_Click()
    Me.partNoTxt.Value = ""
    Me.partRevComBox.Value = ""
    Me.partNameTxt.Value = ""
    Me.partTypComBox.Value = ""
    Me.partStatusTxt.Value = ""
    Me.partGrpTxt.Value = ""
    Me.partSrcTxt.Value = ""
    Me.oldPrtTxt.Value = ""
    Me.wghtTxt.Value = ""
    Me.gradeTxt.Value = ""
End Sub

Private Sub newPartSbmt_Click()
    'Validation
    If Me.partNoTxt.Value = "" Then
        MsgBox "Please enter Part No.", vbCritical
        Exit Sub
    End If
    
    If Me.partRevComBox.Value = "" Then
        MsgBox "Please enter Revision.", vbCritical
        Exit Sub
    End If
    
    If Me.partNameTxt.Value = "" Then
        MsgBox "Please enter Part Name.", vbCritical
        Exit Sub
    End If
    
    If Me.partTypComBox.Value = "" Then
        MsgBox "Please selectr Part Type.", vbCritical
        Exit Sub
    End If
    
    If Me.partStatusTxt.Value = "" Then
        MsgBox "Please enter Part Status.", vbCritical
        Exit Sub
    End If
    
    If Me.partGrpTxt.Value = "" Then
        MsgBox "Please enter Part Group.", vbCritical
        Exit Sub
    End If
    
    If Me.partSrcTxt.Value = "" Then
        MsgBox "Please enter Part Source.", vbCritical
        Exit Sub
    End If
    
    If Me.oldPrtTxt.Value = "" Then
        MsgBox "Please enter Old Part No.", vbCritical
        Exit Sub
    End If
    
'    If Me.wghtTxt.Value = "" Then
'        MsgBox "Weight is auto-calculated. Please be sure it's filled.", vbCritical
'        Exit Sub
'    End If
'     If Me.gradeTxt.Value = "" Then
'        MsgBox "Please enter Grade.", vbCritical
'        Exit Sub
'    End If
    
    If Me.descTxt.Value = "" Then
        MsgBox "Please enter Description.", vbCritical
        Exit Sub
    End If
    
    If Me.descTxt.Value = "" Then
        MsgBox "Please enter Building Code.", vbCritical
        Exit Sub
    End If
''''''''''
    Dim wSht As Worksheet
    Dim n As Long
    Set wSht = ThisWorkbook.Sheets("Decals")
    
    If Application.WorksheetFunction.CountIf(wSht.Range("A:A"), Me.partNoTxt.Value) > 0 Then
        MsgBox "This Part Number exist in spreadsheet", vbCritical
        Exit Sub
    End If
    
    n = wSht.Range("A" & Application.Rows.Count).End(xlUp).Row
    
    wSht.Range("A" & n + 1).Value = Me.partNoTxt.Value
    wSht.Range("B" & n + 1).Value = Me.partRevComBox.Value
    wSht.Range("C" & n + 1).Value = Me.partNameTxt.Value
    wSht.Range("D" & n + 1).Value = Me.partTypComBox.Value
    wSht.Range("E" & n + 1).Value = Me.partStatusTxt.Value
    wSht.Range("F" & n + 1).Value = Me.partGrpTxt.Value
    wSht.Range("G" & n + 1).Value = Me.partSrcTxt.Value
    wSht.Range("I" & n + 1).Value = Me.oldPrtTxt.Value
    wSht.Range("J" & n + 1).Value = Me.wghtTxt.Value
    wSht.Range("L" & n + 1).Value = Me.gradeTxt.Value
    wSht.Range("M" & n + 1).Value = Me.descTxt.Value
    wSht.Range("N" & n + 1).Value = Me.bldCodeTxt.Value
    
''''''''Clear
    
    MsgBox "New Part Added", vbInformation
    
End Sub

Private Sub UserForm_Activate()

    Dim wdthSize As Long, hghtSize As Long
    Dim guagSize As Variant, wghtCalc As Variant, wghtCalcA As Variant, wghtCalcB As Variant, wghtCalcRnd As Variant
'    Dim partRevComBox As Variant
    
'    wdthSize = InputBox("Enter Width.")
'    hghtSize = InputBox("Enter Heigh.")
'    guagSize = InputBox("Enter Guage/Thickness of Metal.")
    partTypComBox.List = Array("Line Marking Signs", "Signs", "Decal/Media", "H41 Marker", _
                        "Wrap Sign Marker", "DRV", "P7 Sign Blanks", "P7 Hardware")
                        
    partRevComBox.List = Array("SCN", "ROL", "EFI")
                        
        
    With Me.wghtTxt
'            wghtCalcA = guagSize * wdthSize
'            wghtCalcB = hghtSize * 0.0977
'            wghtCalc = wghtCalcA * wghtCalcB
'            wghtCalcRnd = WorksheetFunction.Round(wghtCalc, 2)
'        Me.wghtTxt.Value = wghtCalcRnd
    End With

End Sub
