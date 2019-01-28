VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Begin VB.Form FormDataRegistrasiDeposit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Registrasi Deposit"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4185
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormDataRegistrasiDeposit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   4185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID"
         Height          =   270
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nama"
         Height          =   270
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "PIN"
         Height          =   270
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   210
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pass_Web"
         Height          =   270
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Website"
         Height          =   270
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal_Registrasi"
         Height          =   270
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   ":"
         Height          =   270
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   ":"
         Height          =   270
         Left            =   1560
         TabIndex        =   11
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   ":"
         Height          =   270
         Left            =   1560
         TabIndex        =   10
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   ":"
         Height          =   270
         Left            =   1560
         TabIndex        =   9
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   ":"
         Height          =   270
         Left            =   1560
         TabIndex        =   8
         Top             =   2160
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   ":"
         Height          =   270
         Left            =   1560
         TabIndex        =   7
         Top             =   2640
         Width           =   45
      End
      Begin VB.Label LabelID 
         AutoSize        =   -1  'True
         Caption         =   "ID"
         Height          =   270
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   120
      End
      Begin VB.Label LabelNama 
         AutoSize        =   -1  'True
         Caption         =   "Nama"
         Height          =   270
         Left            =   1680
         TabIndex        =   5
         Top             =   720
         Width           =   360
      End
      Begin VB.Label LabelPIN 
         AutoSize        =   -1  'True
         Caption         =   "PIN"
         Height          =   270
         Left            =   1680
         TabIndex        =   4
         Top             =   1200
         Width           =   210
      End
      Begin VB.Label LabelPassWeb 
         AutoSize        =   -1  'True
         Caption         =   "Pass_Web"
         Height          =   270
         Left            =   1680
         TabIndex        =   3
         Top             =   1680
         Width           =   690
      End
      Begin VB.Label LabelWebsite 
         AutoSize        =   -1  'True
         Caption         =   "Website"
         Height          =   270
         Left            =   1680
         TabIndex        =   2
         Top             =   2160
         Width           =   525
      End
      Begin VB.Label LabelTanggalRegistrasi 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal_Registrasi"
         Height          =   270
         Left            =   1680
         TabIndex        =   1
         Top             =   2640
         Width           =   1215
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
   Begin isButton3.isButton cmEdit 
      Height          =   495
      Left            =   2040
      TabIndex        =   19
      Top             =   3240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Icon            =   "FormDataRegistrasiDeposit.frx":000C
      Style           =   7
      Caption         =   "   &Edit"
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
   Begin isButton3.isButton cmTutup 
      Height          =   495
      Left            =   3120
      TabIndex        =   20
      Top             =   3240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Icon            =   "FormDataRegistrasiDeposit.frx":0166
      Style           =   7
      Caption         =   "     &Tutup"
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
Attribute VB_Name = "FormDataRegistrasiDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    Nyambungg
    With AdodcMain
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From TableRegistrasiDeposit"
        .Refresh
        .Refresh
        .Refresh
        .Refresh
    End With
    With Me
        .LabelID.Caption = AdodcMain.Recordset.Fields(0).Value
        .LabelNama.Caption = AdodcMain.Recordset.Fields(1).Value
        .LabelPIN.Caption = AdodcMain.Recordset.Fields(2).Value
        .LabelPassWeb.Caption = AdodcMain.Recordset.Fields(3).Value
        .LabelWebsite.Caption = AdodcMain.Recordset.Fields(4).Value
        .LabelTanggalRegistrasi.Caption = AdodcMain.Recordset.Fields(5).Value
    End With
    RemoveCancelMenuItem Me
End Sub

Private Sub cmEdit_Click()
    X = MsgBox("Anda Yakin ingin merubah data ini?", vbQuestion + vbYesNo)
        If X = vbYes Then
            With FormEditDataRegistrasiDeposit
                .textID.Text = LabelID.Caption
                .textNama.Text = LabelNama.Caption
                .textPIN.Text = LabelPIN.Caption
                .textPassWeb.Text = LabelPassWeb.Caption
                .TextWebsite.Text = LabelWebsite.Caption
                .textTanggalRegistrasi.Text = LabelTanggalRegistrasi.Caption
                Unload Me
                .Show vbModal, Me
            End With
        End If
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub
