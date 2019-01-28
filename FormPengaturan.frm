VERSION 5.00
Object = "{02353968-C1C9-4E0A-88D3-18759BDC60FE}#1.0#0"; "AeroSuite.ocx"
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Begin VB.Form FormPengaturan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pengaturan"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7320
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormPengaturan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   5895
      Begin VB.ComboBox cmbThema 
         Height          =   390
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   4080
         Width           =   1935
      End
      Begin VB.ComboBox cmbDefaultSimpan 
         Height          =   390
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3600
         Width           =   2775
      End
      Begin AeroSuite.AeroCheckBox cekKunciTabel 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         Align           =   0
         Caption         =   "Kunci Tabel"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   0
         MousePointer    =   0
         MouseIcon       =   "FormPengaturan.frx":27A2
         Value           =   1
      End
      Begin AeroSuite.AeroCheckBox cekIzinkanUpdateTabel 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         Align           =   0
         Caption         =   "Izinkan Update Tabel"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   0
         MousePointer    =   0
         MouseIcon       =   "FormPengaturan.frx":27BE
         Value           =   0
      End
      Begin AeroSuite.AeroCheckBox CekToolTipText 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         Align           =   0
         Caption         =   "Aktikan Tips Balon"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   0
         MousePointer    =   0
         MouseIcon       =   "FormPengaturan.frx":27DA
         Value           =   1
      End
      Begin AeroSuite.AeroCheckBox CekStatusBawahFormManage 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         Align           =   0
         Caption         =   "Tampilkan Status Bawah di Jendela Manage"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   0
         MousePointer    =   0
         MouseIcon       =   "FormPengaturan.frx":27F6
         Value           =   0
      End
      Begin AeroSuite.AeroCheckBox CekReset 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         Align           =   0
         Caption         =   "Reset Input Setelah Data Disimpan"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   0
         MousePointer    =   0
         MouseIcon       =   "FormPengaturan.frx":2812
         Value           =   1
      End
      Begin AeroSuite.AeroCheckBox cekPertahankanDataBaru 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         Align           =   0
         Caption         =   "Pertahankan Data Yang Terakhir Disimpan di Jendela Input Data Baru"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   0
         MousePointer    =   0
         MouseIcon       =   "FormPengaturan.frx":282E
         Value           =   1
      End
      Begin AeroSuite.AeroCheckBox CekSetKeHariIni 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         Align           =   0
         Caption         =   "Set Otomatis Tanggal Saat Input Data Baru"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   0
         MousePointer    =   0
         MouseIcon       =   "FormPengaturan.frx":284A
         Value           =   0
      End
      Begin AeroSuite.AeroCheckBox cekGridLines 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         Align           =   0
         Caption         =   "Tampilan gridlines pada tabel depan"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   0
         MousePointer    =   0
         MouseIcon       =   "FormPengaturan.frx":2866
         Value           =   1
      End
      Begin AeroSuite.AeroCheckBox cekSimpanLokasiTerakhirBon 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   3120
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         Align           =   0
         Caption         =   "Simpan Penyimpanan Lokasi Bon Terakhir"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Agency FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   0
         MousePointer    =   0
         MouseIcon       =   "FormPengaturan.frx":2882
         Value           =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Thema :"
         Height          =   270
         Left            =   120
         TabIndex        =   12
         Top             =   4080
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Simpan  :"
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   3600
         Width           =   1065
      End
   End
   Begin XPEngine.XPControl XPEngine 
      Left            =   1080
      Top             =   2760
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin isButton3.isButton cmTutup 
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   4080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Icon            =   "FormPengaturan.frx":289E
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
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButton3.isButton cmOK 
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Icon            =   "FormPengaturan.frx":29F8
      Style           =   6
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
End
Attribute VB_Name = "FormPengaturan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    XPEngine.StartEngine
    With cmbDefaultSimpan
        .Clear
        .AddItem "RikySoft Catatan File (*.rcf)", 0
        .AddItem "Text File (*.txt)", 1
    End With
    With cmbThema
        .Clear
        .AddItem "RS_Default", 0
        .AddItem "RS_Soft", 1
        .AddItem "RS_Flat", 2
        .AddItem "RS_Java", 3
        .AddItem "RS_Office_XP", 4
        .AddItem "RS_Windows_XP", 5
        .AddItem "RS_Windows_Theme", 6
        .AddItem "RS_Plastik", 7
        .AddItem "RS_Galaxy", 8
        .AddItem "RS_Keramik", 9
        .AddItem "RS_MacOSX", 10
    End With
    R = SendMessageLong(cmbDefaultSimpan.hwnd, CB_SETDROPPEDWIDTH, 250, 0)
    R = SendMessageLong(cmbThema.hwnd, CB_SETDROPPEDWIDTH, 200, 0)
    AmbilPengaturan
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
Sub SimpanPengaturan()
    SaveSetting App.Title, "RSBR", Me.cekKunciTabel.Name, Me.cekKunciTabel.Value
    SaveSetting App.Title, "RSBR", Me.cekIzinkanUpdateTabel.Name, Me.cekIzinkanUpdateTabel.Value
    SaveSetting App.Title, "RSBR", Me.CekToolTipText.Name, Me.CekToolTipText.Value
    SaveSetting App.Title, "RSBR", Me.cmbDefaultSimpan.Name, Me.cmbDefaultSimpan.ListIndex
    SaveSetting App.Title, "RSBR", Me.CekStatusBawahFormManage.Name, Me.CekStatusBawahFormManage.Value
    SaveSetting App.Title, "RSBR", Me.CekReset.Name, Me.CekReset.Value
    SaveSetting App.Title, "RSBR", Me.cekPertahankanDataBaru.Name, Me.cekPertahankanDataBaru.Value
    SaveSetting App.Title, "RSBR", Me.cmbThema.Name, Me.cmbThema.ListIndex
    SaveSetting App.Title, "RSBR", Me.CekSetKeHariIni.Name, Me.CekSetKeHariIni.Value
    SaveSetting App.Title, "RSBR", Me.cekGridLines.Name, Me.cekGridLines.Value
End Sub
Sub AmbilPengaturan()
    Me.cekKunciTabel.Value = GetSetting(App.Title, "RSBR", Me.cekKunciTabel.Name, Me.cekKunciTabel.Value)
    Me.cekIzinkanUpdateTabel.Value = GetSetting(App.Title, "RSBR", Me.cekIzinkanUpdateTabel.Name, Me.cekIzinkanUpdateTabel.Value)
    CekToolTipText.Value = GetSetting(App.Title, "RSBR", Me.CekToolTipText.Name, Me.CekToolTipText.Value)
    cmbDefaultSimpan.ListIndex = GetSetting(App.Title, "RSBR", Me.cmbDefaultSimpan.Name, Me.cmbDefaultSimpan.ListIndex)
    CekStatusBawahFormManage.Value = GetSetting(App.Title, "RSBR", Me.CekStatusBawahFormManage.Name, Me.CekStatusBawahFormManage.Value)
    CekReset.Value = GetSetting(App.Title, "RSBR", Me.CekReset.Name, Me.CekReset.Value)
    cekPertahankanDataBaru.Value = GetSetting(App.Title, "RSBR", Me.cekPertahankanDataBaru.Name, Me.cekPertahankanDataBaru.Value)
    cmbThema.ListIndex = GetSetting(App.Title, "RSBR", Me.cmbThema.Name, Me.cmbThema.ListIndex)
    CekSetKeHariIni.Value = GetSetting(App.Title, "RSBR", Me.CekSetKeHariIni.Name, Me.CekSetKeHariIni.Value)
    cekGridLines.Value = GetSetting(App.Title, "RSBR", Me.cekGridLines.Name, Me.cekGridLines.Value)
End Sub

Private Sub cmOK_Click()
    SimpanPengaturan
    FormUtama.AturKontrol
    Unload Me
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub
