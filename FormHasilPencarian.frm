VERSION 5.00
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FormHasilPencarian 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hasil Ditemukan"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6195
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormHasilPencarian.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XPEngine.XPControl XPEngine 
      Left            =   5280
      Top             =   2400
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin VB.TextBox TextDitemukan 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "FormHasilPencarian.frx":014A
      Top             =   120
      Width           =   4935
   End
   Begin isButton3.isButton cmSimpan 
      Height          =   390
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   688
      Icon            =   "FormHasilPencarian.frx":0150
      Style           =   6
      Caption         =   "    &Simpan"
      IconAlign       =   1
      iNonThemeStyle  =   0
      HighlightColor  =   14737632
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButton3.isButton cmCopy 
      Height          =   390
      Left            =   5160
      TabIndex        =   2
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   688
      Icon            =   "FormHasilPencarian.frx":02AA
      Style           =   6
      Caption         =   "&Copy"
      IconAlign       =   1
      iNonThemeStyle  =   0
      HighlightColor  =   14737632
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButton3.isButton cmTutup 
      Height          =   390
      Left            =   5160
      TabIndex        =   3
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   688
      Icon            =   "FormHasilPencarian.frx":0404
      Style           =   6
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FormHasilPencarian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    TextDitemukan.Locked = True
    RemoveCancelMenuItem Me
    XPEngine.StartEngine
    AturThema
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
Private Sub cmCopy_Click()
    TextDitemukan.SelStart = TextDitemukan.SelLength
    Clipboard.Clear
    Clipboard.SetText TextDitemukan.Text
End Sub

Private Sub cmSimpan_Click()
    On Error GoTo ErrorHandler
    CommonDialog1.Filter = "All Files (*.*)|*.*|RikySoft Catatan Files (*.rcf)|*.rcf|Text Files (*.txt)|*.txt"
    DefaultFormat
    CommonDialog1.ShowSave
    CommonDialog1.FileName = CommonDialog1.FileName
    Dim iFile As Integer
    Dim SaveFileFromTB As Boolean
    Dim TxtBox As Object
    Dim FilePath As String
    Dim Append As Boolean
    iFile = FreeFile
    If Append Then
    Open CommonDialog1.FileName For Append As #iFile
    Else
    Open CommonDialog1.FileName For Output As #iFile
    End If
    Print #iFile, TextDitemukan.Text
    SaveFileFromTB = True
ErrorHandler:
    Close #iFile
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub
