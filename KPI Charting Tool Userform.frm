VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KPIChartingTool 
   Caption         =   "UserForm1"
   ClientHeight    =   11370
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   10880
   OleObjectBlob   =   "KPI Charting Tool Userform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "KPIChartingTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cancel_Click()

Unload Me

End Sub

Private Sub clear_Click()

Call UserForm_Initialize

End Sub

Private Sub UserForm_Initialize()

'Initialize text boxes

report_title.Value = ""
improvement1f.Value = ""
improvement2f.Value = ""
improvement3f.Value = ""
improvement4f.Value = ""
improvement5f.Value = ""

'initialize arrays

Dim i As Variant

With mm1f
        For i = 1 To 12
            .AddItem (i)
        Next i
End With


With mm2f
        For i = 1 To 12
            .AddItem (i)
        Next i
End With


With mm3f
        For i = 1 To 12
            .AddItem (i)
        Next i
End With


With mm4f
        For i = 1 To 12
            .AddItem (i)
        Next i
End With


With mm5f
        For i = 1 To 12
            .AddItem (i)
        Next i
End With

Dim d As Integer
d = year(Now)

With yyyy1f
        .AddItem (d - 3)
        .AddItem (d - 2)
        .AddItem (d - 1)
        .AddItem (d)
End With

With yyyy2f
        .AddItem (d - 3)
        .AddItem (d - 2)
        .AddItem (d - 1)
        .AddItem (d)
End With

With yyyy3f
        .AddItem (d - 3)
        .AddItem (d - 2)
        .AddItem (d - 1)
        .AddItem (d)
End With

With yyyy4f
        .AddItem (d - 3)
        .AddItem (d - 2)
        .AddItem (d - 1)
        .AddItem (d)
End With

With yyyy5f
        .AddItem (d - 3)
        .AddItem (d - 2)
        .AddItem (d - 1)
        .AddItem (d)
End With


End Sub

Private Sub create_report_click()

If Len(Me.report_title) > 0 Then
    title = report_title.Value
Else
    title = "New KPI Improvement Chart"
End If

'Improvement 1
If Len(Me.improvement1f) > 0 Then
    improvement1 = improvement1f.Value
    mm1 = mm1f.Value
    yyyy1 = yyyy1f.Value
Else
    improvement1 = ""
    mm1 = 0
    yyyy1 = 0
End If

'Improvement 2
If Len(Me.improvement2f) > 0 Then
    improvement2 = improvement2f.Value
    mm2 = mm2f.Value
    yyyy2 = yyyy2f.Value
Else
    improvement2 = ""
    mm2 = 0
    yyyy2 = 0
End If

'Improvement 3
If Len(Me.improvement3f) > 0 Then
    improvement3 = improvement3f.Value
    mm3 = mm3f.Value
    yyyy3 = yyyy3f.Value
Else
    improvement3 = ""
    mm3 = 0
    yyyy3 = 0
End If



'Improvement 4
If Len(Me.improvement4f) > 0 Then
    improvement4 = improvement4f.Value
    mm4 = mm4f.Value
    yyyy4 = yyyy4f.Value
Else
    improvement4 = ""
    mm4 = 0
    yyyy4 = 0
End If

'Improvement 5
If Len(Me.improvement5f) > 0 Then
    improvement5 = improvement5f.Value
    mm5 = mm5f.Value
    yyyy5 = yyyy5f.Value
Else
    improvement5 = ""
    mm5 = 0
    yyyy5 = 0
End If



Call create_report_now

Unload Me

End Sub
