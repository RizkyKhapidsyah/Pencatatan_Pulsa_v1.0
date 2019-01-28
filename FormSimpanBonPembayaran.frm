VERSION 5.00
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Begin VB.Form FormSimpanBonPembayaran 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simpan Bon"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17535
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormSimpanBonPembayaran.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   17535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   5655
      Begin VB.TextBox TextLokasiPath 
         Height          =   390
         Left            =   1320
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   720
         Width           =   3255
      End
      Begin VB.ComboBox cmbFormatFile 
         Height          =   390
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox TextNamaFile 
         Height          =   390
         Left            =   1320
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   4215
      End
      Begin isButton3.isButton cmTentukan 
         Height          =   375
         Left            =   4680
         TabIndex        =   6
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Style           =   6
         Caption         =   "&Tentukan.."
         IconAlign       =   1
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
      Begin VB.Label LabelFormat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label7"
         Height          =   270
         Left            =   4080
         TabIndex        =   14
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1200
         TabIndex        =   12
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lokasi/Path"
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1200
         TabIndex        =   10
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Format"
         Height          =   270
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama File"
         Height          =   270
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   600
      End
   End
   Begin XPEngine.XPControl XPEngine 
      Left            =   2280
      Top             =   1920
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin isButton3.isButton cmBatal 
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormSimpanBonPembayaran.frx":0442
      Style           =   6
      Caption         =   "&Batal"
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
   Begin isButton3.isButton cmSimpan 
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormSimpanBonPembayaran.frx":059C
      Style           =   6
      Caption         =   "&Simpan"
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
   Begin isButton3.isButton cmDefault 
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormSimpanBonPembayaran.frx":06F6
      Style           =   6
      Caption         =   "&Default"
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
   Begin VB.Label LabelPath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      Height          =   270
      Left            =   0
      TabIndex        =   15
      Top             =   2640
      Width           =   420
   End
End
Attribute VB_Name = "FormSimpanBonPembayaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sHari As String      'Deklarasi variabel global, karena digunakan
Dim aHari                'oleh lebih dari satu prosedur

Sub AturKontrol()
    aHari = Array("Minggu", "Senin", "Selasa", "Rabu", _
                  "Kamis", "Jumat", "Sabtu")
    sHari = aHari(Abs(Weekday(Date) - 1))  'Tampilkan hari
    Kalimat = "" & sHari & ", " _
                     & Format(Date, "dd mmmm yyyy")
    
    XPEngine.StartEngine
    RemoveCancelMenuItem Me
    AturThema
    With Me
        .cmbFormatFile.Clear
        .cmbFormatFile.AddItem "RikySoft Catatan File (*.rcf)", 0
        .cmbFormatFile.AddItem "Microsoft Word Document File (*.doc)", 1
        .cmbFormatFile.AddItem "Text File (*.txt)", 2
        .cmbFormatFile.ListIndex = 0
        .TextLokasiPath.Text = App.Path & "\" & Kalimat
        .TextLokasiPath.Locked = True
        .TextNamaFile = FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value
    End With
    AktifkanToolTipText
    R = SendMessageLong(cmbFormatFile.hwnd, CB_SETDROPPEDWIDTH, 250, 0)
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
Sub AktifkanToolTipText()
    With Me
        .TextNamaFile.ToolTipText = .TextNamaFile.Text
        .TextLokasiPath.ToolTipText = .TextLokasiPath.Text
    End With
End Sub

Private Sub cmBatal_Click()
    Unload Me
End Sub


Private Sub cmDefault_Click()
    With Me
        .TextLokasiPath.Text = App.Path & "\" & Kalimat
        .TextLokasiPath.Locked = True
        .TextNamaFile = FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value
    End With
End Sub

Private Sub cmSimpan_Click()
    Select Case cmbFormatFile.ListIndex
        Case Is = 0
            LokasiFile = TextLokasiPath.Text & "\" & TextNamaFile.Text & ".rcf"
        Case Is = 1
            LokasiFile = TextLokasiPath.Text & "\" & TextNamaFile.Text & ".doc"
        Case Is = 2
            LokasiFile = TextLokasiPath.Text & "\" & TextNamaFile.Text & ".txt"
    End Select
X = FreeFile
Open LokasiFile For Output As #X
Print #X, "==================================" & vbCrLf & _
            FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
            "==================================" & vbCrLf & vbCrLf & _
            FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
            FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
            FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
            FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
            FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
            FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
            FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
            FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
            FormManage.AdodcMain.Recordset.Fields(8).Name & " : " & FormManage.AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
            FormManage.AdodcMain.Recordset.Fields(9).Name & " : " & FormManage.AdodcMain.Recordset.Fields(9).Value & vbCrLf & _
            FormManage.AdodcMain.Recordset.Fields(10).Name & " : " & FormManage.AdodcMain.Recordset.Fields(10).Value & vbCrLf & _
            FormManage.AdodcMain.Recordset.Fields(11).Name & " : " & FormManage.AdodcMain.Recordset.Fields(11).Value & vbCrLf & _
            FormManage.AdodcMain.Recordset.Fields(12).Name & " : " & FormManage.AdodcMain.Recordset.Fields(12).Value & vbCrLf & _
            FormManage.AdodcMain.Recordset.Fields(13).Name & " : " & FormManage.AdodcMain.Recordset.Fields(13).Value & vbCrLf & _
            FormManage.AdodcMain.Recordset.Fields(14).Name & " : " & FormManage.AdodcMain.Recordset.Fields(14).Value & vbCrLf & _
            "=================================="
Close #X
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

Private Sub TextLokasiPath_Change()
    AktifkanToolTipText
End Sub

Private Sub TextNamaFile_Change()
    AktifkanToolTipText
End Sub
