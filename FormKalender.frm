VERSION 5.00
Object = "{62E1564C-1E05-44A7-A921-FD8347F324A5}#1.0#0"; "HookMenu.ocx"
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.ocx"
Begin VB.Form FormKalender 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "-----------------"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   5190
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormKalender.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   960
      Top             =   3600
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   1
      Bmp:1           =   "FormKalender.frx":030A
      Key:1           =   "#menuSHI"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSACAL.Calendar Kalender 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      _Version        =   524288
      _ExtentX        =   9128
      _ExtentY        =   5953
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2014
      Month           =   9
      Day             =   13
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButton3.isButton cmOK_UntukDataBaru 
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   3480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "FormKalender.frx":0732
      Style           =   0
      Caption         =   "&OK"
      IconAlign       =   1
      iNonThemeStyle  =   0
      HighlightColor  =   14737632
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButton3.isButton cmTutup 
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "FormKalender.frx":0B84
      Style           =   0
      Caption         =   "&Tutup"
      IconAlign       =   1
      iNonThemeStyle  =   0
      HighlightColor  =   14737632
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButton3.isButton cmOK_FormUtama 
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   3480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "FormKalender.frx":0CDE
      Style           =   0
      Caption         =   "&OK"
      IconAlign       =   1
      iNonThemeStyle  =   0
      HighlightColor  =   14737632
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LabelView 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-------------"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   780
   End
   Begin VB.Menu menuTool 
      Caption         =   "Tool"
      Begin VB.Menu menuSHI 
         Caption         =   "Set ke Hari Ini"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu menuLKT 
         Caption         =   "Lompat Ke Tanggal.."
      End
   End
End
Attribute VB_Name = "FormKalender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    With Kalender
        .Day = Day(Date)
        .Month = Month(Date)
        .Year = Year(Date)
    End With
    LabelView.Caption = Kalender.Day & " - " & Kalender.Month & " - " & Kalender.Year
    AturThema
    RemoveCancelMenuItem Me
    menuTool.Visible = False
End Sub
Sub AturThema()
    For Each Objek In Me
        If TypeName(Objek) = "isButton" Then
            Select Case FormPengaturan.cmbThema.ListIndex
            Case Is = 0
                Objek.Style = 0
            Case Is = 1
                Objek.Style = 1
            Case Is = 2
                Objek.Style = 2
            Case Is = 3
                Objek.Style = 3
            Case Is = 4
                Objek.Style = 4
            Case Is = 5
                Objek.Style = 5
            Case Is = 6
                Objek.Style = 6
            Case Is = 7
                Objek.Style = 7
            Case Is = 8
                Objek.Style = 8
            Case Is = 9
                Objek.Style = 9
            Case Is = 10
                Objek.Style = 10
            End Select
        End If
    Next
End Sub


Private Sub cmOK_FormUtama_Click()
    Unload Me
End Sub

Private Sub cmOK_UntukDataBaru_Click()
    FormBaru.textTanggalBayar.Text = Me.LabelView.Caption
    Unload Me
    With FormBaru
        .cmbStatusBayar.ListIndex = 0
        .textUangBayar.SetFocus
    End With
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

Private Sub Kalender_Click()
    LabelView.Caption = Kalender.Day & " - " & Kalender.Month & " - " & Kalender.Year
End Sub

Private Sub Kalender_DblClick()
    PopupMenu menuTool
End Sub

Private Sub Kalender_NewMonth()
    LabelView.Caption = Kalender.Day & " - " & Kalender.Month & " - " & Kalender.Year
End Sub

Private Sub Kalender_NewYear()
    LabelView.Caption = Kalender.Day & " - " & Kalender.Month & " - " & Kalender.Year
End Sub

Private Sub menuLKT_Click()
    FormLompatKeTanggal.Show vbModal, Me
End Sub

Private Sub menuSHI_Click()
    With Kalender
        .Day = Day(Date)
        .Month = Month(Date)
        .Year = Year(Date)
    End With
    LabelView.Caption = Kalender.Day & " - " & Kalender.Month & " - " & Kalender.Year
End Sub
