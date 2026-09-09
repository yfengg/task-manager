VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDailyActivities 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   3770
   OleObjectBlob   =   "frmDailyActivities.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDailyActivities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub UserForm_Initialize()
    cmbActivity.Clear
    cmbActivity.AddItem "Study"
    cmbActivity.AddItem "Sleep"
    cmbActivity.AddItem "Work"
    cmbActivity.AddItem "Meal"
    cmbActivity.AddItem "Rest"
    cmbActivity.AddItem "Other"
End Sub
Private Sub btnSubmit_Click()

    Dim ws As Worksheet
    Dim nextRow As Long
    Dim activity As String
    Dim hrs As Double
    Dim actDate As Date
    
    Set ws = ThisWorkbook.Sheets("ActivityLog")
    
    If cmbActivity.Value = "" Then
        MsgBox "Please choose an activity.", vbExclamation
        Exit Sub
    End If
    
    If IsNumeric(txtHours.Value) = False Then
        MsgBox "Hours spent must be a number.", vbExclamation
        Exit Sub
    End If
    
    hrs = CDbl(txtHours.Value)
    
    If IsDate(txtDate.Value) = False Then
        MsgBox "Enter a valid date.", vbExclamation
        Exit Sub
    End If
    
    actDate = CDate(txtDate.Value)

    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = actDate
    ws.Cells(nextRow, 2).Value = cmbActivity.Value
    ws.Cells(nextRow, 3).Value = hrs

    MsgBox "Daily activity logged!", vbInformation
    
    Unload Me

End Sub

