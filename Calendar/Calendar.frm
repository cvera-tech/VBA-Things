VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calendar 
   Caption         =   "Calendar"
   ClientHeight    =   6730
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5280
   OleObjectBlob   =   "Calendar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const DisabledColor As Long = &HD0D0D0
Private Const EnabledColor As Long = &HFFFFFF
Private Const SelectedColor As Long = &HFFFF64
Private EventsFlag As Boolean   ' To prevent multiple events from being fired by field updates
Private ConfirmFlag As Boolean  ' For handling the cancel button
Private CurrentDate As Date     ' Stores the tentative output date

Public Function SelectDate(InputDate As Date) As Date
    ' Set the initial selected date
    Call SetSelectedDateTextBox(InputDate)
    CurrentDate = InputDate
    Call UpdateCalendar(VBA.Month(InputDate), VBA.Year(InputDate))
    
    Call Me.Show
    
    If ConfirmFlag = True Then
        'SelectDate = VBA.DateValue(Me.SelectedDateTextBox.Value)
        SelectDate = CurrentDate
    Else
        SelectDate = InputDate
    End If
    
    Call VBA.Unload(Me)
End Function

Private Sub UserForm_Initialize()
    'Dim CurrentDate As Date
    Dim MonthIndex As Long
    Dim YearIndex As Long
    
    ' Disable events from being handled during initialization
    EventsFlag = False
    
    ' Fill in the month and year combo boxes
    With Me.MonthComboBox
        For MonthIndex = 1 To 12
            .AddItem (VBA.MonthName(MonthIndex))
        Next MonthIndex
    End With
    
    With Me.YearComboBox
        For YearIndex = VBA.Year(VBA.Date) - 20 To VBA.Year(VBA.Date)
            .AddItem (YearIndex)
        Next YearIndex
    End With
    
    ConfirmFlag = False
    
    ' Enable event handling
    EventsFlag = True
    
End Sub

' Attempts to parse SelectedDateTextBox to a Date
' *** Does not trigger any event ***
Private Function ParseSelectedDate() As Date
    Dim DateString As String
    DateString = SelectedDateTextBox.Value
    If VBA.IsDate(DateString) Then
        ParseSelectedDate = VBA.DateValue(DateString)
    Else
        MsgBox ("Unable to parse date. Please enter date in the ""m/d/yyyy"" format.")
        
        ' Reset text box to previous valid date
        'SelectedDateTextBox.Value = VBA.Format(CurrentDate, "m/d/yyyy")
        'ParseSelectedDate = CurrentDate
    End If
End Function

' This procedure changes the month and year comboboxes and calls the
' SetButtons procedure to update the commandbuttons
' *** Triggers events on the ComboBoxes! ***
Private Sub UpdateCalendar(Month As Long, Year As Long)
    Me.MonthComboBox.Value = VBA.MonthName(Month)
    Me.YearComboBox.Value = Year
    Me.CurrentMonthLabel.Caption = VBA.MonthName(Month) & " " & Year
    Call SetSelectedDateTextBox(CurrentDate)
    Call SetButtons(Month, Year)
End Sub

' Updates all the CommandButtons.
' *** Does not trigger any event ***
Private Sub SetButtons(Month As Long, Year As Long)
    Dim ButtonIndex As Long
    Dim Button As MSForms.CommandButton
    Dim FirstDay As Long
    Dim LastDay As Long
    Dim Offset As Long
    Dim NumDays As Long
    
    ' Clear all button captions
    For ButtonIndex = 1 To 42
        Set Button = Me.Controls("CommandButton" & ButtonIndex)
        Button.Caption = ""
    Next ButtonIndex
    
    FirstDay = VBA.DateSerial(Year, Month, 1)
    LastDay = VBA.DateSerial(Year, Month + 1, 1) - 1
    Offset = VBA.Weekday(FirstDay, vbSunday) - 1
    NumDays = LastDay - FirstDay + 1
    
    ' Fill in the correct buttons
    For ButtonIndex = 1 To NumDays
        Set Button = Me.Controls("CommandButton" & ButtonIndex + Offset)
        Button.Caption = ButtonIndex
    Next ButtonIndex
    
    ' Enable the correct buttons
    For ButtonIndex = 1 To 42
        Set Button = Me.Controls("CommandButton" & ButtonIndex)
        
        If Button.Caption = "" Then
            Button.Enabled = False
        Else
            Button.Enabled = True
        End If
    Next ButtonIndex
    
    Call SetButtonColors(Month, Year)
End Sub

' Sets the colors of the buttons
Private Sub SetButtonColors(Month As Long, Year As Long)
    Dim Index As Long
    Dim Button As MSForms.CommandButton
    Dim Offset As Long
    Dim FirstDay As Long
    
    For Index = 1 To 42
        Set Button = Me.Controls("CommandButton" & Index)
        If Button.Caption = "" Then
            Button.BackColor = DisabledColor
        Else
            Button.BackColor = EnabledColor
        End If
    Next Index
    
    FirstDay = VBA.DateSerial(Year, Month, 1)
    Offset = VBA.Weekday(FirstDay, vbSunday) - 1
    If VBA.Month(CurrentDate) = Month And VBA.Year(CurrentDate) = Year Then
        Me.Controls("CommandButton" & (VBA.Day(CurrentDate) + Offset)).BackColor = SelectedColor
    End If
End Sub

Private Sub SetSelectedDateTextBox(NewDate As Date)
    Me.SelectedDateTextBox.Value = VBA.Format(NewDate, "m/d/yyyy")
End Sub

'================
' Event Handlers
'
' The EventsFlag variable is required to prevent these handlers from
' firing multiple times whenever any of the controls are updated.
'
'================

Private Sub MonthComboBox_Click()
    Call UpdateMonth
End Sub

Private Sub MonthComboBox_AfterUpdate()
    Call UpdateMonth
End Sub

Private Sub UpdateMonth()
    If EventsFlag = True Then
        EventsFlag = False
        If Me.YearComboBox.Value <> "" Then
            Dim NewDate As Date
            NewDate = VBA.DateValue(Me.MonthComboBox.Value & " 1, " & Me.YearComboBox.Value)
            Call UpdateCalendar(VBA.Month(NewDate), VBA.Year(NewDate))
        End If
        EventsFlag = True
    End If
End Sub

Private Sub YearComboBox_Click()
    Call UpdateYear
End Sub

Private Sub YearComboBox_AfterUpdate()
    Call UpdateYear
End Sub

Private Sub UpdateYear()
    If EventsFlag = True Then
        EventsFlag = False
        
        If Me.MonthComboBox.Value <> "" Then
            Dim NewDate As Date
            NewDate = VBA.DateValue(Me.MonthComboBox.Value & " 1, " & Me.YearComboBox.Value)
            Call UpdateCalendar(VBA.Month(NewDate), VBA.Year(NewDate))
        End If
        
        EventsFlag = True
    End If
End Sub

Private Sub NextMonthImage_Click()
    Call MonthImageUpdate(1)
End Sub

Private Sub PreviousMonthImage_Click()
    Call MonthImageUpdate(-1)
End Sub

Private Sub MonthImageUpdate(Interval As Long)
    If EventsFlag = True Then
        EventsFlag = False
        
        Dim CurrentMonth As Date
        Dim NextMonth As Date
        CurrentMonth = VBA.DateValue(Me.MonthComboBox.Value & " 1, " & Me.YearComboBox.Value)
        NextMonth = VBA.DateAdd("m", Interval, CurrentMonth)
        Call UpdateCalendar(VBA.Month(NextMonth), VBA.Year(NextMonth))
        
        EventsFlag = True
    End If
End Sub

' Triggers when SelectedDateTextBox loses focus
Private Sub SelectedDateTextBox_AfterUpdate()
    If EventsFlag = True Then
        EventsFlag = False
        
        Dim NewDate As Date
        NewDate = ParseSelectedDate
        If NewDate <> 0 Then
            CurrentDate = NewDate
            Call UpdateCalendar(VBA.Month(NewDate), VBA.Year(NewDate))
        Else
            Call SetSelectedDateTextBox(CurrentDate)
        End If
        
        EventsFlag = True
    End If
End Sub

Private Sub TodayButton_Click()
    CurrentDate = VBA.Date()
    Call SetSelectedDateTextBox(CurrentDate)
    Call UpdateCalendar(VBA.Month(CurrentDate), VBA.Year(CurrentDate))
End Sub

Public Sub ChooseDayButton_Click()
    ConfirmFlag = True
    Call Me.Hide
End Sub

Private Sub CancelButton_Click()
    Call Me.Hide
End Sub

' Sets the text on the SelectedDateTextBox
' *** Triggers events on the TextBox! ***
Private Sub ButtonClick(Button As MSForms.CommandButton)
    If EventsFlag = True Then
        EventsFlag = False
        
        If Button.Caption <> "" Then
            Dim SelectedDateString As String
            Dim SelectedDate As Date
            SelectedDateString = Me.MonthComboBox.Value & " " & Button.Caption & ", " & Me.YearComboBox
            SelectedDate = VBA.DateValue(SelectedDateString)
            CurrentDate = SelectedDate
            Call SetSelectedDateTextBox(SelectedDate)
            Call SetButtonColors(VBA.Month(SelectedDate), VBA.Year(SelectedDate))
        End If
        
        EventsFlag = True
    End If
End Sub

Private Sub CommandButton1_Click()
    Call ButtonClick(Me.CommandButton1)
End Sub
Private Sub CommandButton2_Click()
    Call ButtonClick(Me.CommandButton2)
End Sub
Private Sub CommandButton3_Click()
    Call ButtonClick(Me.CommandButton3)
End Sub
Private Sub CommandButton4_Click()
    Call ButtonClick(Me.CommandButton4)
End Sub
Private Sub CommandButton5_Click()
    Call ButtonClick(Me.CommandButton5)
End Sub
Private Sub CommandButton6_Click()
    Call ButtonClick(Me.CommandButton6)
End Sub
Private Sub CommandButton7_Click()
    Call ButtonClick(Me.CommandButton7)
End Sub
Private Sub CommandButton8_Click()
    Call ButtonClick(Me.CommandButton8)
End Sub
Private Sub CommandButton9_Click()
    Call ButtonClick(Me.CommandButton9)
End Sub
Private Sub CommandButton10_Click()
    Call ButtonClick(Me.CommandButton10)
End Sub
Private Sub CommandButton11_Click()
    Call ButtonClick(Me.CommandButton11)
End Sub
Private Sub CommandButton12_Click()
    Call ButtonClick(Me.CommandButton12)
End Sub
Private Sub CommandButton13_Click()
    Call ButtonClick(Me.CommandButton13)
End Sub
Private Sub CommandButton14_Click()
    Call ButtonClick(Me.CommandButton14)
End Sub
Private Sub CommandButton15_Click()
    Call ButtonClick(Me.CommandButton15)
End Sub
Private Sub CommandButton16_Click()
    Call ButtonClick(Me.CommandButton16)
End Sub
Private Sub CommandButton17_Click()
    Call ButtonClick(Me.CommandButton17)
End Sub
Private Sub CommandButton18_Click()
    Call ButtonClick(Me.CommandButton18)
End Sub
Private Sub CommandButton19_Click()
    Call ButtonClick(Me.CommandButton19)
End Sub
Private Sub CommandButton20_Click()
    Call ButtonClick(Me.CommandButton20)
End Sub
Private Sub CommandButton21_Click()
    Call ButtonClick(Me.CommandButton21)
End Sub
Private Sub CommandButton22_Click()
    Call ButtonClick(Me.CommandButton22)
End Sub
Private Sub CommandButton23_Click()
    Call ButtonClick(Me.CommandButton23)
End Sub
Private Sub CommandButton24_Click()
    Call ButtonClick(Me.CommandButton24)
End Sub
Private Sub CommandButton25_Click()
    Call ButtonClick(Me.CommandButton25)
End Sub
Private Sub CommandButton26_Click()
    Call ButtonClick(Me.CommandButton26)
End Sub
Private Sub CommandButton27_Click()
    Call ButtonClick(Me.CommandButton27)
End Sub
Private Sub CommandButton28_Click()
    Call ButtonClick(Me.CommandButton28)
End Sub
Private Sub CommandButton29_Click()
    Call ButtonClick(Me.CommandButton29)
End Sub
Private Sub CommandButton30_Click()
    Call ButtonClick(Me.CommandButton30)
End Sub
Private Sub CommandButton31_Click()
    Call ButtonClick(Me.CommandButton31)
End Sub
Private Sub CommandButton32_Click()
    Call ButtonClick(Me.CommandButton32)
End Sub
Private Sub CommandButton33_Click()
    Call ButtonClick(Me.CommandButton33)
End Sub
Private Sub CommandButton34_Click()
    Call ButtonClick(Me.CommandButton34)
End Sub
Private Sub CommandButton35_Click()
    Call ButtonClick(Me.CommandButton35)
End Sub
Private Sub CommandButton36_Click()
    Call ButtonClick(Me.CommandButton36)
End Sub
Private Sub CommandButton37_Click()
    Call ButtonClick(Me.CommandButton37)
End Sub
Private Sub CommandButton38_Click()
    Call ButtonClick(Me.CommandButton38)
End Sub
Private Sub CommandButton39_Click()
    Call ButtonClick(Me.CommandButton39)
End Sub
Private Sub CommandButton40_Click()
    Call ButtonClick(Me.CommandButton40)
End Sub
Private Sub CommandButton41_Click()
    Call ButtonClick(Me.CommandButton41)
End Sub
Private Sub CommandButton42_Click()
    Call ButtonClick(Me.CommandButton42)
End Sub
