VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   0
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYOUTRTL = &H400000

Dim WithEvents cmdPrevMonth As CommandButton
Attribute cmdPrevMonth.VB_VarHelpID = -1
Dim WithEvents cmdNextMonth As CommandButton
Attribute cmdNextMonth.VB_VarHelpID = -1
Dim WithEvents cmdCurrentMonth As CommandButton
Attribute cmdCurrentMonth.VB_VarHelpID = -1
Dim WithEvents cmdToggleMonth As CommandButton
Attribute cmdToggleMonth.VB_VarHelpID = -1
Dim startDate As Date
Dim starthebrewDate As hdate

Dim prvCell As TextBox
Dim DayInfo As TextBox

Private Const cell_h = 130
Private Const cell_w = 120

Dim WeekdayLabels(NUM_WEEKDAYS) As Label
Dim CalendarCells(NUM_ROWS, NUM_COLS) As TextBox
Dim CalendarMonthMatrix As Variant
Dim is_hebrew As Boolean
Dim here As location

Private Sub Form_Load()
    With here ' =Yerushalayim
    .latitude = 31.788
    .longitude = 35.218
    .elevation = 800
    End With

    'Me.ScaleMode = vbPixels
    SetFormRTL Me.hWnd
    InitializeCalendar
    startDate = Date
    starthebrewDate = ConvertDate(Date)
    is_hebrew = True
    UpdateCalendar IIf(is_hebrew, starthebrewDate.month, month(startDate)), IIf(is_hebrew, starthebrewDate.year, year(startDate))
End Sub

Private Sub cmdToggleMonth_Click()
    is_hebrew = Not is_hebrew
    cmdToggleMonth.Caption = IIf(is_hebrew, "מעבר לחודש לועזי", "מעבר לחודש עברי")
    UpdateCalendar IIf(is_hebrew, starthebrewDate.month, month(startDate)), IIf(is_hebrew, starthebrewDate.year, year(startDate))
End Sub


Private Sub cmdPrevMonth_Click()
    startDate = DateAdd("m", -1, startDate)
    HDateAddMonth starthebrewDate, -1
    UpdateCalendar IIf(is_hebrew, starthebrewDate.month, month(startDate)), IIf(is_hebrew, starthebrewDate.year, year(startDate))
End Sub

Private Sub cmdCurrentMonth_Click()
    startDate = Date
    starthebrewDate = ConvertDate(Date)
    UpdateCalendar IIf(is_hebrew, starthebrewDate.month, month(startDate)), IIf(is_hebrew, starthebrewDate.year, year(startDate))
End Sub

Private Sub cmdNextMonth_Click()
    startDate = DateAdd("m", 1, startDate)
    HDateAddMonth starthebrewDate, 1
    UpdateCalendar IIf(is_hebrew, starthebrewDate.month, month(startDate)), IIf(is_hebrew, starthebrewDate.year, year(startDate))
End Sub

Private Sub InitializeCalendar()
    Dim I As Integer, J As Integer, Cnt As Integer

    ' Create Weekday labels (column headers)
    For I = 0 To NUM_WEEKDAYS - 1
        Set WeekdayLabels(I) = Controls.Add("VB.Label", "cmdWeekday" & I, Me)
        With WeekdayLabels(I)
            .Caption = WeekdayName(I + 1, True)
            .Left = I * (cell_w * Screen.TwipsPerPixelX)
            .Top = Screen.TwipsPerPixelY * 6
            .Width = (cell_w * Screen.TwipsPerPixelX)
            .Height = (cell_h * Screen.TwipsPerPixelY) / 4
            .Visible = True
            .Alignment = 2
            
        End With
    Next I
        
    ' create buttons (prv/next month, current month, toggle loazi/hebrew calendar)
    Set cmdPrevMonth = Controls.Add("VB.CommandButton", "cmdPrevMonth", Me)
    With cmdPrevMonth
        .Caption = "חודש קודם"
        .Left = (NUM_WEEKDAYS) * (cell_w * Screen.TwipsPerPixelX)
        .Top = 0
        .Width = (cell_w * Screen.TwipsPerPixelX) * (2 / 3)
        .Height = (cell_h * Screen.TwipsPerPixelY) / 4
        .Visible = True
    End With
    Set cmdNextMonth = Controls.Add("VB.CommandButton", "cmdNextMonth", Me)
    With cmdNextMonth
        .Caption = "חודש הבא"
        .Left = (NUM_WEEKDAYS + (2 / 3)) * (cell_w * Screen.TwipsPerPixelX)
        .Top = 0
        .Width = (cell_w * Screen.TwipsPerPixelX) * (2 / 3)
        .Height = (cell_h * Screen.TwipsPerPixelY) / 4
        .Visible = True
    End With
    Set cmdCurrentMonth = Controls.Add("VB.CommandButton", "cmdCurrentMonth", Me)
    With cmdCurrentMonth
        .Caption = "חודש נוכחי"
        .Left = (NUM_WEEKDAYS + (4 / 3)) * (cell_w * Screen.TwipsPerPixelX)
        .Top = 0
        .Width = (cell_w * Screen.TwipsPerPixelX) * (2 / 3)
        .Height = (cell_h * Screen.TwipsPerPixelY) / 4
        .Visible = True
    End With
    Set cmdToggleMonth = Controls.Add("VB.CommandButton", "cmdToggleMonth", Me)
    With cmdToggleMonth
        .Caption = "מעבר לחודש לועזי"
        .Left = (NUM_WEEKDAYS + 2) * (cell_w * Screen.TwipsPerPixelX)
        .Top = 0
        .Width = (cell_w * Screen.TwipsPerPixelX)
        .Height = (cell_h * Screen.TwipsPerPixelY) / 4
        .Visible = True
    End With
    
    ' Create Calendar cells
    For I = 0 To NUM_ROWS - 1
        For J = 0 To NUM_COLS - 1
            If Cnt > 0 Then Load Text2(Cnt)
            Set CalendarCells(I, J) = Text2(Cnt) 'Controls.Add("VB.TextBox", "txtDate" & I & J, Me)
            Cnt = Cnt + 1
            With CalendarCells(I, J)
                .Top = (I) * (cell_h * Screen.TwipsPerPixelY) + ((cell_h * Screen.TwipsPerPixelY) / 4)
                .Left = J * (cell_w * Screen.TwipsPerPixelX)
                .Width = (cell_w * Screen.TwipsPerPixelX)
                .Height = (cell_h * Screen.TwipsPerPixelY)
                .BorderStyle = 1 ' Remove border
                .FontName = "Arial"
                .FontSize = 7.5
                .Locked = True ' Make read-only
                .Visible = True
                .Alignment = 2
                .Font.Charset = 177
            End With
        Next J
    Next I
    Set DayInfo = Text1 'Controls.Add("VB.TextBox", "txtDayInfo", Me)
    With DayInfo
        .Top = ((cell_h * Screen.TwipsPerPixelY) / 4)
        .Left = (NUM_ROWS + 1) * (cell_w * Screen.TwipsPerPixelX)
        .Width = (cell_w * Screen.TwipsPerPixelX) * 3
        .Height = (cell_h * Screen.TwipsPerPixelY) * NUM_ROWS
        .BorderStyle = 1 ' Remove border
        .FontName = "Arial"
        .FontSize = 13
        .Locked = False ' Make read-only
        .Visible = True
        .Alignment = 2
        .BackColor = vbInfoBackground
    End With
    Me.Height = NUM_ROWS * (cell_h * Screen.TwipsPerPixelY) + (cell_h * Screen.TwipsPerPixelY) / 4 + (Me.Height - Me.ScaleHeight)
    Me.Width = NUM_COLS * (cell_w * Screen.TwipsPerPixelX) + (cell_w * Screen.TwipsPerPixelX) * 3 + (Me.Width - Me.ScaleWidth)
End Sub

Private Sub UpdateCalendar(ByVal month_in As Integer, ByVal year_in As Integer)
    Dim currentDay As Date
    Dim I As Integer, J As Integer
    Dim row As Integer, col As Integer
    Dim vDay() As String
    Dim starthdate As hdate
    
    'load month dates array, as table of 7 days X 6 weeks. Each item is in this string format:
    '<date>;<belongs to month?[0/1];<first date of month>
    CalendarMonthMatrix = calendar_utils_get_month_matrix(month_in, year_in, is_hebrew)
        
    For I = 0 To NUM_ROWS - 1
        For J = 0 To NUM_COLS - 1
                        'update the cells according to the table data
            vDay = Split(CalendarMonthMatrix(I, J), ";")
            CalendarCells(I, J).Text = calendar_utils_get_day_info(CDate(vDay(0)))
            CalendarCells(I, J).Tag = CalendarMonthMatrix(I, J) 'keep day data in the cell's tag
            CalendarCells(I, J).BackColor = IIf(CDate(vDay(0)) = Date, vbYellow, IIf(vDay(1) = "1", vbWindowBackground, vbButtonFace)) 'yellow for today, white for day of the month, gray for days from prv/nxt months
            CalendarCells(I, J).ForeColor = SystemColorConstants.vbWindowText 'IIf(vDay(1) = "1", SystemColorConstants.vbWindowText, SystemColorConstants.vbGrayText)
            CalendarCells(I, J).FontBold = IIf(vDay(1) = "1", True, False)
        Next J
    Next I
    
    'load molad info
    startDate = CDate(vDay(2)) 'DateSerial(year_in, month_in, 1)
    starthdate = ConvertDate(startDate)
    DayInfo.Text = vbCrLf & "בחר יום מהלוח כדי להציג זמנים" & vbCrLf & vbCrLf & calendar_utils_get_month_molad_info(startDate)

    'show selected month in the title bar
    If is_hebrew = True Then
        Me.Caption = NumToHMonth(month_in, starthdate.leap) & " " & NumToHChar(year_in)
    Else
        Me.Caption = MonthName(month_in) & " " & year_in
    End If
End Sub

Private Sub Text2_Click(Index As Integer)
    Dim vDay() As String
    
    'restore previous cell to non-selected color
    If Not prvCell Is Nothing Then
        vDay = Split(prvCell.Tag, ";")
        If Not prvCell Is Text2(Index) Then prvCell.BackColor = IIf(vDay(0) = Date, vbYellow, IIf(vDay(1) = "1", vbWindowBackground, vbButtonFace))
    End If
        
    Set prvCell = Text2(Index)
    'get day data from tag of selected cell
    vDay = Split(prvCell.Tag, ";")
    prvCell.BackColor = SystemColorConstants.vbHighlight 'highlight selected cell
        'show zmanin & limud info
    DayInfo.Text = vbCrLf & "-זמני היום-" & vbCrLf & calendar_utils_get_zmanim_info(CDate(vDay(0)), here) & vbCrLf & vbCrLf & "-לימוד יומי-" & vbCrLf & calendar_utils_get_limud_info(CDate(vDay(0)))
End Sub


Public Sub SetFormRTL(ByVal hWnd As Long)
    Dim dwExStyle As Long
    
    ' Get the current extended window style
    dwExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    
    ' Add WS_EX_LAYOUTRTL to the extended window style
    dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
    
    ' Set the modified extended window style
    SetWindowLong hWnd, GWL_EXSTYLE, dwExStyle
End Sub
