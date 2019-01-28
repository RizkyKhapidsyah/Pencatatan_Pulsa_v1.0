VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Begin VB.Form FormEditDataRegistrasiDeposit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Data"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4170
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormEditDataRegistrasiDeposit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin XPEngine.XPControl XPEngine 
      Left            =   600
      Top             =   3360
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.TextBox textTanggalRegistrasi 
         Height          =   390
         Left            =   1680
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   2640
         Width           =   2175
      End
      Begin VB.TextBox TextWebsite 
         Height          =   390
         Left            =   1680
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox textPassWeb 
         Height          =   390
         Left            =   1680
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox textPIN 
         Height          =   390
         Left            =   1680
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox textNama 
         Height          =   390
         Left            =   1680
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox textID 
         Height          =   390
         Left            =   1680
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   ":"
         Height          =   270
         Left            =   1560
         TabIndex        =   12
         Top             =   2640
         Width           =   45
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   ":"
         Height          =   270
         Left            =   1560
         TabIndex        =   11
         Top             =   2160
         Width           =   45
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   ":"
         Height          =   270
         Left            =   1560
         TabIndex        =   10
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   ":"
         Height          =   270
         Left            =   1560
         TabIndex        =   9
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   ":"
         Height          =   270
         Left            =   1560
         TabIndex        =   8
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   ":"
         Height          =   270
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal_Registrasi"
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Website"
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pass_Web"
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "PIN"
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nama"
         Height          =   270
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID"
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   120
      End
   End
   Begin MSAdodcLib.Adodc AdodcMain 
      Height          =   330
      Left            =   6600
      Top             =   2520
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin isButton3.isButton cmSimpan 
      Height          =   495
      Left            =   2040
      TabIndex        =   13
      Top             =   3240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Icon            =   "FormEditDataRegistrasiDeposit.frx":000C
      Style           =   7
      Caption         =   "      &Simpan"
      IconAlign       =   1
      iNonThemeStyle  =   0
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
   Begin isButton3.isButton cmBatal 
      Height          =   495
      Left            =   3120
      TabIndex        =   14
      Top             =   3240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Icon            =   "FormEditDataRegistrasiDeposit.frx":0166
      Style           =   7
      Caption         =   "     &Batal"
      IconAlign       =   1
      iNonThemeStyle  =   0
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
End
Attribute VB_Name = "FormEditDataRegistrasiDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    XPEngine.StartEngine
    KosongkanTextBox
    With Me
        .textID.Text = FormDataRegistrasiDeposit.LabelID.Caption
        .textNama.Text = FormDataRegistrasiDeposit.LabelNama.Caption
        .textPIN.Text = FormDataRegistrasiDeposit.LabelPIN.Caption
        .textPassWeb.Text = FormDataRegistrasiDeposit.LabelPassWeb.Caption
        .TextWebsite.Text = FormDataRegistrasiDeposit.LabelWebsite.Caption
        .textTanggalRegistrasi.Text = FormDataRegistrasiDeposit.LabelTanggalRegistrasi.Caption
        .textID.MaxLength = 254
        .textNama.MaxLength = 254
        .textPIN.MaxLength = 254
        .textPassWeb.MaxLength = 254
        .TextWebsite.MaxLength = 254
        .textTanggalRegistrasi.MaxLength = 254
    End With
    RemoveCancelMenuItem Me
End Sub

Sub KosongkanTextBox()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            Objek.Text = ""
        End If
    Next
End Sub

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmSimpan_Click()
    If textID.Text = "" Then
        MsgBox "Silahkan isi data ID Anda!", vbExclamation + vbOKOnly, ""
        textID.SetFocus
    ElseIf textNama.Text = "" Then
        MsgBox "Silahkan isi data Nama Anda!", vbExclamation + vbOKOnly, ""
        textNama.SetFocus
    ElseIf textPIN.Text = "" Then
        MsgBox "Silahkan isi data PIN Anda!", vbExclamation + vbOKOnly, ""
        textPIN.SetFocus
    ElseIf textPassWeb.Text = "" Then
        MsgBox "Silahkan isi data Password di WEB!", vbExclamation + vbOKOnly, ""
        textPassWeb.SetFocus
    ElseIf TextWebsite.Text = "" Then
        MsgBox "Silahkan isi data Website!", vbExclamation + vbOKOnly, ""
        TextWebsite.SetFocus
    ElseIf textTanggalRegistrasi.Text = "" Then
        MsgBox "Silahkan isi data tanggal registrasi Anda!", vbExclamation + vbOKOnly, ""
        textTanggalRegistrasi.SetFocus
    Else
        X = MsgBox("Apakah Anda yakin ingin menyimpan data ini?", vbQuestion + vbYesNo)
        If X = vbYes Then
            With FormDataRegistrasiDeposit.AdodcMain
                .Recordset.Delete
                .Recordset.AddNew
                .Recordset.Fields(0).Value = textID.Text
                .Recordset.Fields(1).Value = textNama.Text
                .Recordset.Fields(2).Value = textPIN.Text
                .Recordset.Fields(3).Value = textPassWeb.Text
                .Recordset.Fields(4).Value = TextWebsite.Text
                .Recordset.Fields(5).Value = textTanggalRegistrasi.Text
                .Recordset.Update
                .Refresh
                .Refresh
            End With
            MsgBox "Data Berhasil diperbarui!", vbInformation + vbOKOnly, ""
            Unload Me
            Unload FormDataRegistrasiDeposit
        End If
    End If
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub
