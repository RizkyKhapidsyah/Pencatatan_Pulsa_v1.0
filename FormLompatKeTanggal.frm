VERSION 5.00
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Begin VB.Form FormLompatKeTanggal 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lompat Ke Tanggal.."
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2775
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   2775
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbTahun 
      Height          =   375
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox cmbBulan 
      Height          =   375
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox cmbTanggal 
      Height          =   375
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin XPEngine.XPControl XPEngine 
      Left            =   840
      Top             =   2520
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin isButton3.isButton cmOK 
      Height          =   390
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   688
      Icon            =   "FormLompatKeTanggal.frx":0000
      Style           =   0
      Caption         =   "&OK"
      IconAlign       =   1
      CaptionAlign    =   2
      iNonThemeStyle  =   0
      HighlightColor  =   14737632
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButton3.isButton cmBatal 
      Height          =   390
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   688
      Icon            =   "FormLompatKeTanggal.frx":0452
      Style           =   0
      Caption         =   "&Batal"
      IconAlign       =   1
      CaptionAlign    =   2
      iNonThemeStyle  =   0
      HighlightColor  =   14737632
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FormLompatKeTanggal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    cmbTanggal.Clear
    For X = 31 To 1 Step -1
        With cmbTanggal
            .AddItem X, 0
            .ListIndex = 0
        End With
    Next X
    cmbBulan.Clear
    For Y = 12 To 1 Step -1
        With cmbBulan
            .AddItem Y, 0
            .ListIndex = 0
        End With
    Next Y
    cmbTahun.Clear
    For Z = 2200 To 1800 Step -1
        With cmbTahun
            .AddItem Z, 0
            .ListIndex = 0
        End With
    Next Z
    With Me
        .cmbTanggal.Text = Day(Date)
        .cmbBulan.Text = Month(Date)
        .cmbTahun.Text = Year(Date)
    End With
    XPEngine.StartEngine
    AturThema
    RemoveCancelMenuItem Me
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

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmOK_Click()
    With FormKalender.Kalender
        .Day = cmbTanggal.Text
        .Month = cmbBulan.Text
        .Year = cmbTahun.Text
    End With
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub
