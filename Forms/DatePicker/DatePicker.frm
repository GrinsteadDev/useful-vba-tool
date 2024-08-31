VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DatePicker 
   Caption         =   "UserForm1"
   ClientHeight    =   7725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6255
   OleObjectBlob   =   "DatePicker.frx":0000
End
Attribute VB_Name = "DatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' DatePicker v2.1.0
'   Designed to replace the Windows 7 DatePicker UserForm Control
' Useage
'
' Dependencies
'   Microsoft Froms 2.0 Object Library
'   https://github.com/GrinsteadDev/useful-vba-tools/blob/main/Forms/FormEvents/LabelEvent.cls

Option Explicit

Private Const DAY_ROW_COUNT = 7
Private Const DAY_COL_COUNT = 7
Private Const DAY_LBL_COUNT = DAY_ROW_COUNT * DAY_COL_COUNT
Private Const DAY_LBL_PADDING = 3
Private Const MONTH_ROW_COUNT = 3
Private Const MONTH_COL_COUNT = 4
Private Const MONTH_LBL_COUNT = MONTH_ROW_COUNT * MONTH_COL_COUNT
Private Const MONTH_LBL_PADDING = 6
Private Const YEAR_ROW_COUNT = 3
Private Const YEAR_COL_COUNT = 3
Private Const YEAR_LBL_COUNT = YEAR_ROW_COUNT * YEAR_COL_COUNT
Private Const YEAR_LBL_PADDING = 6

Private Const HIGHLIGHT_BG = vbInactiveTitleBar
Private Const HIGHLIGHT_FG = vbInactiveCaptionText

Private Const CURRENT_BG = vbHighlight
Private Const CURRENT_FG = vbHighlightText

Private Const DEFAULT_BG = vbButtonFace
Private Const DEFAULT_FG = vbButtonText

Private Const NOT_FOCUS_BG = vb3DLight
Private Const NOT_FOCUS_FG = vbGrayText

Private LblCollection As Collection
Private LastHover As MSForms.Label
Private LastHoverBg As Long
Private LastHoverFg As Long
Private CurrTar As MSForms.Label
Private CurrDate As Date
Private IsCanceled As Boolean

'' Helper
Private Function IsEven(Val As Long) As Boolean
    IsEven = Not CBool(Val And 1)
End Function

Private Sub Center()
    With Me
        .Left = Application.Left + (Application.Width / 2) - (.Width / 2)
        .Top = Application.Top + ((Application.Height + (Application.Height - Application.UsableHeight)) / 2) - .Height
    End With
End Sub

Private Sub AddLabels(RowCount As Long, ColCount As Long, Total As Long, PageFrame As MSForms.Frame, Padding As Long)
    Dim i As Long, lbl_obj As MSForms.Label, _
        lbl_event As LabelEvent, lbl_width As Double, _
        lbl_height As Double, lbl_left As Double, _
        lbl_top As Double
    
    lbl_width = PageFrame.Width / ColCount
    lbl_height = PageFrame.Height / RowCount
    For i = 1 To Total
        lbl_left = (i Mod ColCount)
        If lbl_left = 0 Then lbl_left = ColCount
        lbl_left = lbl_left * lbl_width - lbl_width
        
        lbl_top = VBA.Int((i - 1) / ColCount) * lbl_height
        
        Set lbl_obj = PageFrame.Controls.Add("Forms.Label.1")
            lbl_obj.Width = lbl_width - (Padding * 2)
            lbl_obj.Height = lbl_height - (Padding * 2)
            lbl_obj.Left = lbl_left + Padding
            lbl_obj.Top = lbl_top + Padding
            lbl_obj.BorderStyle = fmBorderStyleSingle
            lbl_obj.TextAlign = fmTextAlignCenter
            
        Set lbl_event = New LabelEvent
        Set lbl_event.FormLabel = lbl_obj
            lbl_event.AddEvent "Lbl_Click", LabelEventTypes.Click, Me
            lbl_event.AddEvent "Lbl_MouseMove", LabelEventTypes.MouseMove, Me
            lbl_event.AddEvent "Day_DblClick", LabelEventTypes.DblClick, Me
            
        
        LblCollection.Add Array(lbl_obj, lbl_event)
        
        Set lbl_obj = Nothing
        Set lbl_event = Nothing
    Next
End Sub

Private Sub ShowDay()
    Me.DayFrame.Visible = True
    Me.MonthFrame.Visible = False
    Me.YearFrame.Visible = False
    Me.DateCaption.Caption = VBA.MonthName(VBA.Month(CurrDate)) & " " & CStr(VBA.Year(CurrDate))
    Me.Caption = VBA.Format(CurrDate, "dddd, mmmm dd")
    
    Dim ft_day As Long, ln_day As Long, _
        ln_day_ln_month As Long, i As Long
    
    ln_day_ln_month = VBA.Day(VBA.DateSerial(VBA.Year(CurrDate), VBA.Month(CurrDate), 0))
    ft_day = VBA.Weekday(VBA.DateSerial(VBA.Year(CurrDate), VBA.Month(CurrDate), 1), vbUseSystemDayOfWeek)
    ln_day = VBA.Day(VBA.DateSerial(VBA.Year(CurrDate), VBA.Month(CurrDate) + 1, 0))
    
    For i = 1 To DAY_COL_COUNT
        With Me.DayFrame.Controls(i - 1)
            .Caption = VBA.WeekdayName(i, True)
            .BackColor = DEFAULT_BG
            .ForeColor = DEFAULT_FG
        End With
    Next
    For i = 1 To ft_day - 1
        With Me.DayFrame.Controls(i - 1 + DAY_COL_COUNT)
            .Caption = CStr(ln_day_ln_month - ft_day + i + 1)
            .BackColor = NOT_FOCUS_BG
            .ForeColor = NOT_FOCUS_FG
            .Tag = -1
        End With
    Next
    For i = ft_day To ft_day + ln_day - 1
        With Me.DayFrame.Controls(i - 1 + DAY_COL_COUNT)
            .Caption = CStr(i - ft_day + 1)
            .BackColor = DEFAULT_BG
            .ForeColor = DEFAULT_FG
            .Tag = 0
            
            If i - ft_day + 1 = VBA.Day(CurrDate) Then
                .BackColor = CURRENT_BG
                .ForeColor = CURRENT_FG
                
                Set CurrTar = Me.DayFrame.Controls(i - 1 + DAY_COL_COUNT)
            End If
        End With
    Next
    For i = ft_day + ln_day + DAY_COL_COUNT To DAY_LBL_COUNT
        With Me.DayFrame.Controls(i - 1)
            .Caption = CStr(i - (ft_day + ln_day + DAY_COL_COUNT) + 1)
            .BackColor = NOT_FOCUS_BG
            .ForeColor = NOT_FOCUS_FG
            .Tag = 1
        End With
    Next
End Sub

Private Sub ShowMonth()
    Me.DayFrame.Visible = False
    Me.MonthFrame.Visible = True
    Me.YearFrame.Visible = False
    Me.DateCaption.Caption = CStr(VBA.Year(CurrDate))
    
    Dim i As Long
    
    For i = 1 To MONTH_LBL_COUNT
        With Me.MonthFrame.Controls(i - 1)
            .Caption = VBA.MonthName(i, True)
            .BackColor = DEFAULT_BG
            .ForeColor = DEFAULT_FG
            
            If i = VBA.Month(CurrDate) Then
                .BackColor = CURRENT_BG
                .ForeColor = CURRENT_FG
                
                Set CurrTar = Me.MonthFrame.Controls(i - 1)
            End If
        End With
    Next
End Sub

Private Sub ShowYear()
    Me.DayFrame.Visible = False
    Me.MonthFrame.Visible = False
    Me.YearFrame.Visible = True
    
    Dim st_year As Long, en_year As Long, _
        i As Long, middle As Long
    
    middle = VBA.Int(YEAR_LBL_COUNT / 2)
    
    If Not IsEven(YEAR_LBL_COUNT) Then middle = middle + 1
    
    Me.DateCaption.Caption = CStr(VBA.Year(CurrDate) - middle - 1) & " - " & CStr(VBA.Year(CurrDate) - middle + YEAR_LBL_COUNT)
    
    For i = 1 To middle - 1
        With Me.YearFrame.Controls(i - 1)
            .Caption = CStr(VBA.Year(CurrDate) - (middle - i))
            .BackColor = DEFAULT_BG
            .ForeColor = DEFAULT_FG
        End With
    Next i
    
    With Me.YearFrame.Controls(middle - 1)
        .Caption = CStr(VBA.Year(CurrDate))
        .BackColor = CURRENT_BG
        .ForeColor = CURRENT_FG
                
        Set CurrTar = Me.YearFrame.Controls(middle - 1)
    End With
    
    For i = middle + 1 To YEAR_LBL_COUNT
        With Me.YearFrame.Controls(i - 1)
            .Caption = CStr(VBA.Year(CurrDate) - (middle - i))
            .BackColor = DEFAULT_BG
            .ForeColor = DEFAULT_FG
        End With
    Next
End Sub

'' Form Events
Private Sub DateCaption_Click()
    If Not Me.YearFrame.Visible Then
        If Me.MonthFrame.Visible Then
            ShowYear
        End If
        
        If Me.DayFrame.Visible Then
            ShowMonth
        End If
    End If
End Sub

Private Sub DateCaption_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not (LastHover Is Nothing Or LastHover Is CurrTar) Then
        With LastHover
            .BackColor = LastHoverBg
            .ForeColor = LastHoverFg
        End With
    End If
    
    Set LastHover = Me.DateCaption
        LastHoverBg = Me.DateCaption.BackColor
        LastHoverFg = Me.DateCaption.ForeColor
    
    Me.DateCaption.ForeColor = vbActiveTitleBar
End Sub

Private Sub Increment_Click()
    If Me.DayFrame.Visible Then
        CurrDate = VBA.DateAdd("m", 1, CurrDate)
        ShowDay
    ElseIf Me.MonthFrame.Visible Then
        CurrDate = VBA.DateAdd("yyyy", 1, CurrDate)
        ShowMonth
    ElseIf Me.YearFrame.Visible Then
        CurrDate = VBA.DateAdd("yyyy", YEAR_LBL_COUNT, CurrDate)
        ShowYear
    End If
End Sub

Private Sub Decrement_Click()
    If Me.DayFrame.Visible Then
        CurrDate = VBA.DateAdd("m", -1, CurrDate)
        ShowDay
    ElseIf Me.MonthFrame.Visible Then
        CurrDate = VBA.DateAdd("yyyy", -1, CurrDate)
        ShowMonth
    ElseIf Me.YearFrame.Visible Then
        CurrDate = VBA.DateAdd("yyyy", -YEAR_LBL_COUNT, CurrDate)
        ShowYear
    End If
End Sub

Private Sub TodayTag_Click()
    CurrDate = VBA.Date
    
    If Me.DayFrame.Visible Then
        ShowDay
    ElseIf Me.MonthFrame.Visible Then
        ShowMonth
    ElseIf Me.YearFrame.Visible Then
        ShowYear
    End If
End Sub

Private Sub TodayTag_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not (LastHover Is Nothing Or LastHover Is CurrTar) Then
        With LastHover
            .BackColor = LastHoverBg
            .ForeColor = LastHoverFg
        End With
    End If
    
    Set LastHover = Me.TodayTag
        LastHoverBg = Me.TodayTag.BackColor
        LastHoverFg = Me.TodayTag.ForeColor
    
    Me.TodayTag.ForeColor = vbActiveTitleBar
End Sub

Private Sub UserForm_Activate()
    Center
    ShowDay
End Sub

Private Sub UserForm_Initialize()
    Dim i As Long
    
    Set LblCollection = New Collection
    CurrDate = VBA.Date
    Me.TodayTag.Caption = "Today " & VBA.Format(VBA.Date, "mm-dd-yyyy")
    
    AddLabels DAY_ROW_COUNT, DAY_COL_COUNT, DAY_LBL_COUNT, Me.DayFrame, DAY_LBL_PADDING
    AddLabels MONTH_ROW_COUNT, MONTH_COL_COUNT, MONTH_LBL_COUNT, Me.MonthFrame, MONTH_LBL_PADDING
    AddLabels YEAR_ROW_COUNT, YEAR_COL_COUNT, YEAR_LBL_COUNT, Me.YearFrame, YEAR_LBL_PADDING
    
    For i = 1 To DAY_COL_COUNT
        LblCollection.Remove 1
    Next
    
    ShowDay
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not (LastHover Is Nothing Or LastHover Is CurrTar) Then
        With LastHover
            .BackColor = LastHoverBg
            .ForeColor = LastHoverFg
        End With
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    IsCanceled = True
End Sub

Private Sub DayFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    UserForm_MouseMove Button, Shift, X, Y
End Sub

Private Sub MonthFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    UserForm_MouseMove Button, Shift, X, Y
End Sub

Private Sub YearFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    UserForm_MouseMove Button, Shift, X, Y
End Sub

Private Sub OkButton_Click()
    IsCanceled = False
    Me.Hide
End Sub

Private Sub CancelButton_Click()
    IsCanceled = True
    Me.Hide
End Sub

'' Dynamic Events
Public Sub Lbl_MouseMove( _
    Sender As MSForms.Label, ByVal Button As Integer, ByVal Shift As Integer, _
    ByVal X As Single, ByVal Y As Single _
)
    If Not (LastHover Is Nothing Or LastHover Is CurrTar) Then
        With LastHover
            .BackColor = LastHoverBg
            .ForeColor = LastHoverFg
        End With
    End If
    
    Set LastHover = Sender
        LastHoverBg = Sender.BackColor
        LastHoverFg = Sender.ForeColor
    
    If Not Sender Is CurrTar Then
        With Sender
            .BackColor = HIGHLIGHT_BG
            .ForeColor = HIGHLIGHT_FG
        End With
    End If
End Sub

Public Sub Lbl_Click(Sender As MSForms.Label)
    If Sender Is CurrTar Then Exit Sub
    
    With Sender
        Select Case .Parent.Name
            Case Me.DayFrame.Name
                CurrDate = VBA.DateValue(CStr(VBA.Month(CurrDate) + CLng(Sender.Tag)) & "/" & Sender.Caption & "/" & CStr(VBA.Year(CurrDate)))
                
                If CLng(Sender.Tag) = 0 Then
                    CurrTar.BackColor = DEFAULT_BG
                    CurrTar.ForeColor = DEFAULT_FG
                    .BackColor = CURRENT_BG
                    .ForeColor = CURRENT_FG
                    
                    Set CurrTar = Sender
                Else
                    ShowDay
                End If
            Case Me.MonthFrame.Name
                CurrDate = VBA.DateValue(Sender.Caption & " " & CStr(VBA.Day(CurrDate)) & ", " & CStr(VBA.Year(CurrDate)))
                ShowDay
            Case Me.YearFrame.Name
                CurrDate = VBA.DateValue(CStr(VBA.Month(CurrDate)) & "/" & CStr(VBA.Day(CurrDate)) & "/" & Sender.Caption)
                ShowMonth
        End Select
    End With
End Sub

Public Sub Day_DblClick(Sender As MSForms.Label, ByVal Cancel As MSForms.ReturnBoolean)
    If Sender.Parent.Name = Me.DayFrame.Name Then
        OkButton_Click
    End If
End Sub

'' Public Properties
Public Property Get SelectedDate() As Date
    SelectedDate = CurrDate
End Property
Public Property Get SelectedYear() As Long
    SelectedYear = VBA.Year(CurrDate)
End Property
Public Property Get SelectedMonth() As Long
    SelectedMonth = VBA.Month(CurrDate)
End Property
Public Property Get Canceled() As Boolean
    Canceled = IsCanceled
End Property

'' Public Methods
Public Function GetDate(Optional DefaultDate As Date) As Date
    Dim d As Date
    
    If d = DefaultDate Then DefaultDate = VBA.Date
    
    CurrDate = DefaultDate
    Me.Show
    
    GetDate = CurrDate
End Function
