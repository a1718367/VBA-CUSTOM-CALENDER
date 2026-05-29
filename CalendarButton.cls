VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalendarButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Expose the button object with events enabled
Public WithEvents ClickableButton As MSForms.CommandButton
Attribute ClickableButton.VB_VarHelpID = -1
' Store the actual date associated with this specific button
'Public ButtonDate As Date
Private ButtonDate As Date
' Reference back to the main form to call its closing logic
Public ParentForm As Object

Private Sub ClickableButton_Click()
    ' Pass the clicked date back to the main form variable
    'ParentForm.selectedDate = ThisButtonDate
    ParentForm.selectedDate = ThisButtonDate
    'frmNew.tbDOB.Value = Format(ThisButtonDate, "Short Date")

    MsgBox "You selected: " & Format(ThisButtonDate, "Long Date")
    
    'ParentForm.Hide
End Sub

' Helper property to access the date safely
Public Property Let ThisButtonDate(d As Date)
    ButtonDate = d
End Property

Public Property Get ThisButtonDate() As Date
    ThisButtonDate = ButtonDate
End Property


