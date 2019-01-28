VERSION 5.00
Object = "{62E1564C-1E05-44A7-A921-FD8347F324A5}#1.0#0"; "HookMenu.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FormBaru 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catatan Baru"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormBaru.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   4800
      Top             =   720
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   1
      Bmp:1           =   "FormBaru.frx":014A
      Key:1           =   "#menuCetakBon"
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
   Begin MSAdodcLib.Adodc AdodcMain 
      Height          =   330
      Left            =   120
      Top             =   960
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
   Begin VB.Frame Frame1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   6855
      Begin VB.ComboBox CmbRefHargaJual 
         Height          =   390
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   3600
         Width           =   975
      End
      Begin VB.ComboBox cmbRefHargaServer 
         Height          =   390
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   3120
         Width           =   975
      End
      Begin VB.ComboBox cmbRefJumlahPulsaDibeli 
         Height          =   390
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox textTanggalBayar 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   1920
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   4560
         Width           =   3255
      End
      Begin VB.TextBox textUangBayar 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   1920
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   4080
         Width           =   3255
      End
      Begin VB.ComboBox cmbJenisProvider 
         Height          =   390
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   2160
         Width           =   1815
      End
      Begin VB.ComboBox cmbProvider 
         Height          =   390
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1680
         Width           =   1215
      End
      Begin VB.ComboBox cmbTanggal 
         Height          =   390
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cmbBulan 
         Height          =   390
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cmbTahun 
         Height          =   390
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox textPenerimaPulsa 
         Height          =   390
         Left            =   1920
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox TextNamaPenerima 
         Height          =   390
         Left            =   1920
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1200
         Width           =   4815
      End
      Begin VB.TextBox textJumlahPulsaDibeli 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   1920
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox textHargaServer 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   1920
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox TextHargaJual 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   1920
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox textLaba 
         Height          =   390
         Left            =   1920
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   5040
         Width           =   1575
      End
      Begin VB.ComboBox cmbStatusBayar 
         Height          =   390
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   5520
         Width           =   1695
      End
      Begin VB.ComboBox cmbHari 
         Height          =   390
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin isButton3.isButton cmSetKeHariIni 
         Height          =   390
         Left            =   5400
         TabIndex        =   12
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   688
         Icon            =   "FormBaru.frx":0572
         Style           =   0
         Caption         =   "Set ke Hari Ini"
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
      Begin VB.ComboBox CMBDataLalu1 
         Height          =   390
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   720
         Width           =   3255
      End
      Begin VB.ComboBox CMBDataLalu2 
         Height          =   390
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   1200
         Width           =   3975
      End
      Begin VB.ComboBox CMBDataLalu3 
         Height          =   390
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   2640
         Width           =   2895
      End
      Begin VB.ComboBox CMBDataLalu4 
         Height          =   390
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   3120
         Width           =   2655
      End
      Begin VB.ComboBox CMBDataLalu5 
         Height          =   390
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   3600
         Width           =   2895
      End
      Begin VB.ComboBox CMBDataLalu6 
         Height          =   390
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   4080
         Width           =   2895
      End
      Begin isButton3.isButton cmListTelepon 
         Height          =   390
         Left            =   5400
         TabIndex        =   51
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   688
         Icon            =   "FormBaru.frx":058E
         Style           =   0
         Caption         =   "List Telepon"
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
      Begin isButton3.isButton cmSet 
         Height          =   390
         Left            =   5280
         TabIndex        =   52
         Top             =   4560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   688
         Style           =   0
         Caption         =   "&Set"
         IconSize        =   10
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
      Begin isButton3.isButton cmHutang 
         Height          =   390
         Left            =   5280
         TabIndex        =   53
         Top             =   4080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   688
         Icon            =   "FormBaru.frx":05AA
         Style           =   0
         Caption         =   "&Hutang"
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
      Begin isButton3.isButton cmTambahRefJumlahPulsaDibeli 
         Height          =   390
         Left            =   6360
         TabIndex        =   60
         Top             =   2640
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   688
         Icon            =   "FormBaru.frx":05C6
         Style           =   0
         Caption         =   "+"
         IconAlign       =   1
         iNonThemeStyle  =   0
         HighlightColor  =   14737632
         Object.ToolTipText     =   "Klik untuk menambah data referensi"
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
      Begin isButton3.isButton cmbTambahRefHargaServer 
         Height          =   390
         Left            =   6360
         TabIndex        =   61
         Top             =   3120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   688
         Icon            =   "FormBaru.frx":05E2
         Style           =   0
         Caption         =   "+"
         IconAlign       =   1
         iNonThemeStyle  =   0
         HighlightColor  =   14737632
         Object.ToolTipText     =   "Klik untuk menambah data referensi"
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
      Begin isButton3.isButton cmbTambahRefHargaJual 
         Height          =   390
         Left            =   6360
         TabIndex        =   62
         Top             =   3600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   688
         Icon            =   "FormBaru.frx":05FE
         Style           =   0
         Caption         =   "+"
         IconAlign       =   1
         iNonThemeStyle  =   0
         HighlightColor  =   14737632
         Object.ToolTipText     =   "Klik untuk menambah data referensi"
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
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref :"
         Height          =   270
         Left            =   4920
         TabIndex        =   56
         Top             =   3600
         Width           =   285
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref :"
         Height          =   270
         Left            =   4920
         TabIndex        =   55
         Top             =   3120
         Width           =   285
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref :"
         Height          =   270
         Left            =   4920
         TabIndex        =   54
         Top             =   2640
         Width           =   285
      End
      Begin VB.Label LabelInfo10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Bayar"
         Height          =   270
         Left            =   120
         TabIndex        =   40
         Top             =   4560
         Width           =   945
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   39
         Top             =   4560
         Width           =   45
      End
      Begin VB.Label LabelInfo9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Uang_Bayar (Rp)"
         Height          =   270
         Left            =   120
         TabIndex        =   38
         Top             =   4080
         Width           =   1125
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   37
         Top             =   4080
         Width           =   45
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   34
         Top             =   2160
         Width           =   45
      End
      Begin VB.Label LabelInfo5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Provider"
         Height          =   270
         Left            =   120
         TabIndex        =   33
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   31
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label LabelInfo4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Provider"
         Height          =   270
         Left            =   120
         TabIndex        =   30
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label LabelInfo1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   270
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   27
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   26
         Top             =   720
         Width           =   45
      End
      Begin VB.Label LabelInfo2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penerima_Pulsa"
         Height          =   270
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   24
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label LabelInfo3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama_Penerima"
         Height          =   270
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   22
         Top             =   2640
         Width           =   45
      End
      Begin VB.Label LabelInfo6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah_Pulsa_Dibeli"
         Height          =   270
         Left            =   120
         TabIndex        =   21
         Top             =   2640
         Width           =   1290
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   20
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label LabelInfo7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga_Server (Rp)"
         Height          =   270
         Left            =   120
         TabIndex        =   19
         Top             =   3120
         Width           =   1275
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   18
         Top             =   3600
         Width           =   45
      End
      Begin VB.Label LabelInfo8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga_Jual (Rp)"
         Height          =   270
         Left            =   120
         TabIndex        =   17
         Top             =   3600
         Width           =   1065
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   16
         Top             =   5040
         Width           =   45
      End
      Begin VB.Label LabelInfo11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Laba (Rp)"
         Height          =   270
         Left            =   120
         TabIndex        =   15
         Top             =   5040
         Width           =   630
      End
      Begin VB.Label LabelInfo12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status_Bayar"
         Height          =   270
         Left            =   120
         TabIndex        =   14
         Top             =   5520
         Width           =   885
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1800
         TabIndex        =   13
         Top             =   4560
         Width           =   45
      End
   End
   Begin XPEngine.XPControl XPEngine 
      Left            =   3000
      Top             =   8040
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin isButton3.isButton cmSimpan 
      Height          =   495
      Left            =   120
      TabIndex        =   47
      Top             =   7440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormBaru.frx":061A
      Style           =   0
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
   Begin isButton3.isButton cmVerifikasi 
      Height          =   495
      Left            =   1440
      TabIndex        =   48
      Top             =   7440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormBaru.frx":0774
      Style           =   0
      Caption         =   "     &Verifikasi"
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
   Begin isButton3.isButton cmReset 
      Height          =   495
      Left            =   2760
      TabIndex        =   49
      Top             =   7440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormBaru.frx":0BC6
      Style           =   0
      Caption         =   "&Reset"
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
      Left            =   5760
      TabIndex        =   50
      Top             =   7440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormBaru.frx":0D20
      Style           =   0
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
   Begin MSAdodcLib.Adodc AdodcCmbRefJumlahPulsaDibeli 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc AdodcCmbRefHargaServer 
      Height          =   330
      Left            =   1200
      Top             =   0
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
   Begin MSAdodcLib.Adodc AdodcCmbRefHargaJual 
      Height          =   330
      Left            =   2400
      Top             =   0
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
   Begin VB.Menu MenuMenu 
      Caption         =   "Menu"
      Begin VB.Menu menuCetakBon 
         Caption         =   "Cetak Bon"
      End
   End
End
Attribute VB_Name = "FormBaru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sHari As String      'Deklarasi variabel global, karena digunakan
Dim aHari                'oleh lebih dari satu prosedur


Sub AturKontrol()
On Error GoTo HancurkanError
    Nyambungg
    With AdodcMain
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From TablePulsa order by Tahun asc, Bulan, Tanggal"
        .Refresh
    End With
    With AdodcCmbRefJumlahPulsaDibeli
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From TableReferensiJumlahPulsaDibeli order by Jumlah_Pulsa_Dibeli asc"
        .Refresh
    End With
    With AdodcCmbRefHargaServer
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From TableReferensiHargaServer order by Harga_Server asc"
        .Refresh
    End With
    With AdodcCmbRefHargaJual
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From TableReferensiHargaJual order by Harga_Jual asc"
        .Refresh
    End With
    With cmbHari
        .Clear
        .AddItem "Senin", 0
        .AddItem "Selasa", 1
        .AddItem "Rabu", 2
        .AddItem "Kamis", 3
        .AddItem "Jum'at", 4
        .AddItem "Sabtu", 5
        .AddItem "Minggu", 6
        .ListIndex = 0
    End With
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
    With cmbStatusBayar
        .Clear
        .AddItem "Lunas", 0
        .AddItem "Hutang", 1
        .AddItem "Menunggu", 2
        .ListIndex = 0
    End With
    With Me
        .textJumlahPulsaDibeli.Alignment = 2
        .textHargaServer.Alignment = 1
        .TextHargaJual.Alignment = 1
        .textLaba.Alignment = 1
    End With
    With cmbProvider
        .Clear
        .AddItem "XL", 0
        .AddItem "Telkomsel", 1
        .AddItem "Indosat", 2
        .AddItem "3 (Tri)", 3
        .AddItem "(CDMA)", 4
        .ListIndex = 0
    End With
    Reset
    Me.Picture = LoadPicture(App.Path & "\image\BannerDataBaru.jpg")
    If FormPengaturan.CekSetKeHariIni.Value = Checked Then cmSetKeHariIni_Click
    IsiCMBDataLalu
    AturThema
    RemoveCancelMenuItem Me
    XPEngine.StartEngine
    textJumlahPulsaDibeli.Text = ""
    textHargaServer.Text = ""
    TextHargaJual.Text = ""
Exit Sub
HancurkanError:
    PusatError
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
Sub Reset()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .Text = ""
                .MaxLength = 254
            End With
        End If
    Next
    With Me
        .textLaba.Text = Val(TextHargaJual.Text) - Val(textHargaServer.Text)
        .textLaba.Locked = True
        .textTanggalBayar.Locked = True
    End With
End Sub
Sub IsiCMBDataLalu()
    With Me
        .CMBDataLalu1.Clear
        .CMBDataLalu2.Clear
        .CMBDataLalu3.Clear
        .CMBDataLalu4.Clear
        .CMBDataLalu5.Clear
        .CMBDataLalu6.Clear
        .cmbRefJumlahPulsaDibeli.Clear
        .cmbRefHargaServer.Clear
        .CmbRefHargaJual.Clear
        Do Until AdodcMain.Recordset.EOF
            .CMBDataLalu1.AddItem AdodcMain.Recordset.Fields(4).Value
            .CMBDataLalu2.AddItem AdodcMain.Recordset.Fields(5).Value
            .CMBDataLalu3.AddItem AdodcMain.Recordset.Fields(8).Value
            .CMBDataLalu4.AddItem AdodcMain.Recordset.Fields(9).Value
            .CMBDataLalu5.AddItem AdodcMain.Recordset.Fields(10).Value
            .CMBDataLalu6.AddItem AdodcMain.Recordset.Fields(11).Value
            AdodcMain.Recordset.MoveNext
        Loop
        AdodcMain.Refresh
        Do Until AdodcCmbRefJumlahPulsaDibeli.Recordset.EOF
            .cmbRefJumlahPulsaDibeli.AddItem AdodcCmbRefJumlahPulsaDibeli.Recordset.Fields(0).Value
            AdodcCmbRefJumlahPulsaDibeli.Recordset.MoveNext
        Loop
        AdodcCmbRefJumlahPulsaDibeli.Refresh
        Do Until AdodcCmbRefHargaServer.Recordset.EOF
            .cmbRefHargaServer.AddItem AdodcCmbRefHargaServer.Recordset.Fields(0).Value
            AdodcCmbRefHargaServer.Recordset.MoveNext
        Loop
        AdodcCmbRefHargaServer.Refresh
        Do Until AdodcCmbRefHargaJual.Recordset.EOF
            .CmbRefHargaJual.AddItem AdodcCmbRefHargaJual.Recordset.Fields(0).Value
            AdodcCmbRefHargaJual.Recordset.MoveNext
        Loop
        AdodcCmbRefHargaJual.Refresh
        .cmbRefJumlahPulsaDibeli.ListIndex = 0
        .cmbRefHargaServer.ListIndex = 0
        .CmbRefHargaJual.ListIndex = 0
    End With
End Sub

Private Sub CMBDataLalu1_Click()
    textPenerimaPulsa.Text = CMBDataLalu1.Text
End Sub

Private Sub CMBDataLalu2_Click()
    TextNamaPenerima.Text = CMBDataLalu2.Text
End Sub

Private Sub CMBDataLalu3_Click()
    textJumlahPulsaDibeli.Text = CMBDataLalu3.Text
End Sub

Private Sub CMBDataLalu4_Click()
    textHargaServer.Text = CMBDataLalu4.Text
End Sub

Private Sub CMBDataLalu5_Click()
    TextHargaJual.Text = CMBDataLalu5.Text
End Sub

Private Sub CMBDataLalu6_Click()
    textUangBayar.Text = CMBDataLalu6.Text
End Sub

Private Sub CmbRefHargaJual_Click()
    TextHargaJual.Text = CmbRefHargaJual.Text
End Sub

Private Sub cmbRefHargaServer_Click()
    textHargaServer.Text = cmbRefHargaServer.Text
End Sub

Private Sub cmbRefJumlahPulsaDibeli_Click()
    textJumlahPulsaDibeli.Text = cmbRefJumlahPulsaDibeli.Text
End Sub

Private Sub cmbTambahRefHargaJual_Click()
    With FormTambahReferensi
        .Caption = "Tambah Referensi (" & Me.LabelInfo8.Caption & ")"
        .LabelKeterangan.Caption = "Tambah Referensi (" & Me.LabelInfo8.Caption & ")"
        .LabelPenentu.Caption = "3"
        .TextKeterangan.Text = ""
        .Show vbModal, Me
    End With
End Sub

Private Sub cmbTambahRefHargaServer_Click()
    With FormTambahReferensi
        .Caption = "Tambah Referensi (" & Me.LabelInfo7.Caption & ")"
        .LabelKeterangan.Caption = "Tambah Referensi (" & Me.LabelInfo7.Caption & ")"
        .LabelPenentu.Caption = "2"
        .TextKeterangan.Text = ""
        .Show vbModal, Me
    End With
End Sub

Private Sub cmHutang_Click()
    textUangBayar.Text = "-"
    textTanggalBayar.Text = "Belum Bayar"
    With cmbStatusBayar
        .ListIndex = 1
        .SetFocus
    End With
End Sub

Private Sub cmbProvider_Click()
    Select Case cmbProvider.ListIndex
        Case Is = 0
            With cmbJenisProvider
                .Clear
                .AddItem "XL", 0
                .AddItem "Axis", 1
                .ListIndex = 0
            End With
        Case Is = 1
            With cmbJenisProvider
                .Clear
                .AddItem "simPATI", 0
                .AddItem "Kartu AS", 1
                .ListIndex = 0
            End With
        Case Is = 2
            With cmbJenisProvider
                .Clear
                .AddItem "IM3", 0
                .AddItem "Matrix", 1
                .AddItem "Mentari", 2
                .ListIndex = 0
            End With
        Case Is = 3
            With cmbJenisProvider
                .Clear
                .AddItem "3 (Tri)", 0
                .ListIndex = 0
            End With
        Case Is = 4
            With cmbJenisProvider
                .Clear
                .AddItem "Smart", 0
                .AddItem "Fren", 1
                .AddItem "SmartFren", 2
                .AddItem "Esia", 3
                .AddItem "Flexi", 4
                .ListIndex = 0
            End With
    End Select
End Sub

Private Sub cmListTelepon_Click()
    With FormListTelepon
        .cmMasukan.Enabled = True
        .cmTambah.Enabled = False
        .cmEdit.Enabled = False
        .cmCari.Enabled = True
        .cmHapus.Enabled = False
        .cmRefresh.Enabled = False
        .Show vbModal, Me
    End With
End Sub

Private Sub cmReset_Click()
    AturKontrol
End Sub

Private Sub cmSet_Click()
    With FormKalender
        .Caption = "Set Tanggal"
        .cmOK_FormUtama.Value = False
        .cmOK_UntukDataBaru.Value = True
        .cmTutup.Visible = True
        .Show vbModal, Me
    End With
End Sub

Public Sub cmSetKeHariIni_Click()
    aHari = Array("Minggu", "Senin", "Selasa", "Rabu", _
                  "Kamis", "Jumat", "Sabtu")
    sHari = aHari(Abs(Weekday(Date) - 1))  'Tampilkan hari
    Select Case sHari
    Case "Minggu"
        cmbHari.ListIndex = 6
    Case "Senin"
        cmbHari.ListIndex = 0
    Case "Selasa"
        cmbHari.ListIndex = 1
    Case "Rabu"
        cmbHari.ListIndex = 2
    Case "Kamis"
        cmbHari.ListIndex = 3
    Case "Jumat"
        cmbHari.ListIndex = 4
    Case "Sabtu"
        cmbHari.ListIndex = 5
    End Select
    cmbTanggal.Text = Day(Date)
    cmbBulan.Text = Month(Date)
    cmbTahun.Text = Year(Date)
End Sub

Private Sub cmSimpan_Click()
    If textPenerimaPulsa.Text = "" Then
        MsgBox "Silahkan isi Nomor yang akan diisikan pulsa!", vbExclamation + vbOKOnly, ""
        textPenerimaPulsa.SetFocus
    ElseIf TextNamaPenerima.Text = "" Then
        MsgBox "Silahkan isi Nama Penerima pulsa!", vbExclamation + vbOKOnly, ""
        TextNamaPenerima.SetFocus
    ElseIf textJumlahPulsaDibeli.Text = "" Then
        MsgBox "Silahkan isi Jumlah pulsa yang akan dibeli!", vbExclamation + vbOKOnly, ""
        textJumlahPulsaDibeli.SetFocus
    ElseIf textHargaServer.Text = "" Then
        MsgBox "Silahkan isi Harga pulsa awal dari server!", vbExclamation + vbOKOnly, ""
        textHargaServer.SetFocus
    ElseIf TextHargaJual.Text = "" Then
        MsgBox "Silahkan isi Harga jual pulsa!", vbExclamation + vbOKOnly, ""
        TextHargaJual.SetFocus
    ElseIf textUangBayar.Text = "" Then
        MsgBox "Silahkan isi jumlah uang pembayaran pulsa!", vbExclamation + vbOKOnly, ""
        textUangBayar.SetFocus
    ElseIf textTanggalBayar.Text = "" Then
        MsgBox "Silahkan isi tanggal bayar", vbExclamation + vbOKOnly, ""
        textTanggalBayar.SetFocus
    Else
        Select Case cmSimpan.Caption
        Case "&Simpan"
            X = MsgBox("Anda yakin ingin menyimpan data ini?", vbQuestion + vbYesNo, "Konfirmasi")
            If X = vbYes Then
                With FormUtama
                    .AdodcMain.Recordset.AddNew
                    .AdodcMain.Recordset.Fields(0).Value = cmbHari.Text
                    .AdodcMain.Recordset.Fields(1).Value = cmbTanggal.Text
                    .AdodcMain.Recordset.Fields(2).Value = cmbBulan.Text
                    .AdodcMain.Recordset.Fields(3).Value = cmbTahun.Text
                    .AdodcMain.Recordset.Fields(4).Value = textPenerimaPulsa.Text
                    .AdodcMain.Recordset.Fields(5).Value = TextNamaPenerima.Text
                    .AdodcMain.Recordset.Fields(6).Value = cmbProvider.Text
                    .AdodcMain.Recordset.Fields(7).Value = cmbJenisProvider.Text
                    .AdodcMain.Recordset.Fields(8).Value = textJumlahPulsaDibeli.Text
                    .AdodcMain.Recordset.Fields(9).Value = textHargaServer.Text
                    .AdodcMain.Recordset.Fields(10).Value = TextHargaJual.Text
                    .AdodcMain.Recordset.Fields(11).Value = textUangBayar.Text
                    .AdodcMain.Recordset.Fields(12).Value = textTanggalBayar.Text
                    .AdodcMain.Recordset.Fields(13).Value = textLaba.Text
                    .AdodcMain.Recordset.Fields(14).Value = cmbStatusBayar.Text
                    .AdodcMain.Recordset.Update
                    .AdodcMain.Refresh
                    .AturKontrol
                End With
                If FormPengaturan.CekReset.Value = Checked Then Reset
                textPenerimaPulsa.SetFocus
                IsiCMBDataLalu
                cmTutup.Caption = "&Tutup"
            End If
        Case "&Update"
            X = MsgBox("Anda yakin ingin memperbarui data ini?", vbQuestion + vbYesNo, "Konfirmasi")
            If X = vbYes Then
                With FormManage
                    .AdodcMain.Recordset.Delete
                    .AdodcMain.Recordset.AddNew
                    .AdodcMain.Recordset.Fields(0).Value = cmbHari.Text
                    .AdodcMain.Recordset.Fields(1).Value = cmbTanggal.Text
                    .AdodcMain.Recordset.Fields(2).Value = cmbBulan.Text
                    .AdodcMain.Recordset.Fields(3).Value = cmbTahun.Text
                    .AdodcMain.Recordset.Fields(4).Value = textPenerimaPulsa.Text
                    .AdodcMain.Recordset.Fields(5).Value = TextNamaPenerima.Text
                    .AdodcMain.Recordset.Fields(6).Value = cmbProvider.Text
                    .AdodcMain.Recordset.Fields(7).Value = cmbJenisProvider.Text
                    .AdodcMain.Recordset.Fields(8).Value = textJumlahPulsaDibeli.Text
                    .AdodcMain.Recordset.Fields(9).Value = textHargaServer.Text
                    .AdodcMain.Recordset.Fields(10).Value = TextHargaJual.Text
                    .AdodcMain.Recordset.Fields(11).Value = textUangBayar.Text
                    .AdodcMain.Recordset.Fields(12).Value = textTanggalBayar.Text
                    .AdodcMain.Recordset.Fields(13).Value = textLaba.Text
                    .AdodcMain.Recordset.Fields(14).Value = cmbStatusBayar.Text
                    .AdodcMain.Recordset.Update
                    .AdodcMain.Refresh
                    .AdodcMain.Refresh
                    .AturUkuranDatagrid
                End With
                FormUtama.AturKontrol
                FormManage.cmRefresh_Click
                Unload Me
            End If
        End Select
    End If
End Sub

Private Sub cmTambahRefJumlahPulsaDibeli_Click()
    With FormTambahReferensi
        .Caption = "Tambah Referensi (" & Me.LabelInfo6.Caption & ")"
        .LabelKeterangan.Caption = "Tambah Referensi (" & Me.LabelInfo6.Caption & ")"
        .LabelPenentu.Caption = "1"
        .TextKeterangan.Text = ""
        .Show vbModal, Me
    End With
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub cmVerifikasi_Click()
    If textPenerimaPulsa.Text = "" Then
        MsgBox "Data Tidak dapat diverifikasi!" & vbCrLf & _
                "Silahkan isi Nomor yang akan diisikan pulsa!", vbExclamation + vbOKOnly, ""
        textPenerimaPulsa.SetFocus
    ElseIf TextNamaPenerima.Text = "" Then
        MsgBox "Data Tidak dapat diverifikasi!" & vbCrLf & _
                "Silahkan isi Nama Penerima pulsa!", vbExclamation + vbOKOnly, ""
        TextNamaPenerima.SetFocus
    ElseIf textJumlahPulsaDibeli.Text = "" Then
        MsgBox "Data Tidak dapat diverifikasi!" & vbCrLf & _
                "Silahkan isi Jumlah pulsa yang akan dibeli!", vbExclamation + vbOKOnly, ""
        textJumlahPulsaDibeli.SetFocus
    ElseIf textHargaServer.Text = "" Then
        MsgBox "Data Tidak dapat diverifikasi!" & vbCrLf & _
                "Silahkan isi Harga pulsa awal dari server!", vbExclamation + vbOKOnly, ""
        textHargaServer.SetFocus
    ElseIf TextHargaJual.Text = "" Then
        MsgBox "Data Tidak dapat diverifikasi!" & vbCrLf & _
                "Silahkan isi Harga jual pulsa!", vbExclamation + vbOKOnly, ""
        TextHargaJual.SetFocus
    ElseIf textUangBayar.Text = "" Then
        MsgBox "Data Tidak dapat diverifikasi!" & vbCrLf & _
                "Silahkan isi jumlah uang pembayaran pulsa!", vbExclamation + vbOKOnly, ""
        textUangBayar.SetFocus
    ElseIf textTanggalBayar.Text = "" Then
        MsgBox "Data Tidak dapat diverifikasi!" & vbCrLf & _
                "Silahkan isi tanggal bayar", vbExclamation + vbOKOnly, ""
        textTanggalBayar.SetFocus
    Else
        MsgBox "Verifikasi OK!" & vbCrLf & _
                "Data sudah bisa disimpan!", vbInformation + vbOKOnly, "Verifikasi OK"
        cmSimpan.SetFocus
    End If
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

Private Sub menuCetakBon_Click()
    On Error GoTo ErrorHandler
        If textPenerimaPulsa.Text = "" Then
            MsgBox "Bon tidak dapat dicetak. Silahkan isi Nomor yang akan diisikan pulsa!", vbExclamation + vbOKOnly, ""
            textPenerimaPulsa.SetFocus
        ElseIf TextNamaPenerima.Text = "" Then
            MsgBox "Bon tidak dapat dicetak. Silahkan isi Nama Penerima pulsa!", vbExclamation + vbOKOnly, ""
            TextNamaPenerima.SetFocus
        ElseIf textJumlahPulsaDibeli.Text = "" Then
            MsgBox "Bon tidak dapat dicetak. Silahkan isi Jumlah pulsa yang akan dibeli!", vbExclamation + vbOKOnly, ""
            textJumlahPulsaDibeli.SetFocus
        ElseIf textHargaServer.Text = "" Then
            MsgBox "Bon tidak dapat dicetak. Silahkan isi Harga pulsa awal dari server!", vbExclamation + vbOKOnly, ""
            textHargaServer.SetFocus
        ElseIf TextHargaJual.Text = "" Then
            MsgBox "Bon tidak dapat dicetak. Silahkan isi Harga jual pulsa!", vbExclamation + vbOKOnly, ""
            TextHargaJual.SetFocus
        ElseIf textUangBayar.Text = "" Then
            MsgBox "Bon tidak dapat dicetak. Silahkan isi jumlah uang pembayaran pulsa!", vbExclamation + vbOKOnly, ""
            textUangBayar.SetFocus
        ElseIf textTanggalBayar.Text = "" Then
            MsgBox "Bon tidak dapat dicetak. Silahkan isi tanggal bayar", vbExclamation + vbOKOnly, ""
            textTanggalBayar.SetFocus
        ElseIf cmbStatusBayar.ListIndex = 1 Or cmbStatusBayar.ListIndex = 2 Then
            MsgBox "Bon tidak dapat dicetak karena status pembayaran masih '" & cmbStatusBayar.Text & "'!", vbExclamation + vbOKOnly, ""
            cmbStatusBayar.SetFocus
        Else
            On Error GoTo ErrorHandler
            CommonDialog1.DialogTitle = "Cetak Bon"
            CommonDialog1.FileName = Me.TextNamaPenerima.Text & " (" & Me.textPenerimaPulsa.Text & ") - " & Me.cmbHari.Text & ", " & Me.cmbTanggal.Text & " - " & Me.cmbBulan.Text & " - " & Me.cmbTahun.Text
            CommonDialog1.Filter = "All Files (*.*)|*.*|RikySoft Catatan Files (*.rcf)|*.rcf|Text Files (*.txt)|*.txt"
            DefaultFormat
            CommonDialog1.ShowSave
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
            Print #iFile, "======================================================"
            Print #iFile, Me.TextNamaPenerima.Text & " (" & Me.textPenerimaPulsa.Text & ") - " & Me.cmbHari.Text & ", " & Me.cmbTanggal.Text & " - " & Me.cmbBulan.Text & " - " & Me.cmbTahun.Text
            Print #iFile, "======================================================"
            Print #iFile, LabelInfo1.Caption & " : " & cmbHari.Text & ", " & cmbTanggal.Text & " - " & cmbBulan.Text & " - " & cmbTahun.Text
            Print #iFile, LabelInfo2.Caption & " : " & textPenerimaPulsa.Text
            Print #iFile, LabelInfo3.Caption & " : " & TextNamaPenerima.Text
            Print #iFile, LabelInfo4.Caption & " : " & cmbProvider.Text
            Print #iFile, LabelInfo5.Caption & " : " & cmbJenisProvider.Text
            Print #iFile, LabelInfo6.Caption & " : " & textJumlahPulsaDibeli.Text
            Print #iFile, LabelInfo7.Caption & " : " & textHargaServer.Text
            Print #iFile, LabelInfo8.Caption & " : " & TextHargaJual.Text
            Print #iFile, LabelInfo9.Caption & " : " & textUangBayar.Text
            Print #iFile, LabelInfo10.Caption & " : " & textTanggalBayar.Text
            Print #iFile, LabelInfo11.Caption & " : " & textLaba.Text
            Print #iFile, LabelInfo12.Caption & " : " & cmbStatusBayar.Text
            Print #iFile, "======================================================"
            SaveFileFromTB = True
                Close #iFile
        End If
ErrorHandler:
    Close #iFile
End Sub

Private Sub TextHargaJual_Change()
    With textLaba
        .Text = Val(TextHargaJual.Text) - Val(textHargaServer.Text)
        .Locked = True
    End With
End Sub

Private Sub TextHargaJual_DblClick()
R = SendMessageLong(CMBDataLalu5.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub TextHargaJual_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbKeySpace) Then
        Beep
        KeyAscii = 0
   End If
End Sub

Private Sub textHargaServer_Change()
    With textLaba
        .Text = Val(TextHargaJual.Text) - Val(textHargaServer.Text)
        .Locked = True
    End With
End Sub

Private Sub textHargaServer_DblClick()
R = SendMessageLong(CMBDataLalu4.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textHargaServer_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbKeySpace) Then
        Beep
        KeyAscii = 0
   End If
End Sub

Private Sub textJumlahPulsaDibeli_DblClick()
    R = SendMessageLong(CMBDataLalu3.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textJumlahPulsaDibeli_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbKeySpace) Then
        Beep
        KeyAscii = 0
   End If
End Sub



Private Sub TextNamaPenerima_DblClick()
    R = SendMessageLong(CMBDataLalu2.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textPenerimaPulsa_DblClick()
    R = SendMessageLong(CMBDataLalu1.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

Private Sub textPenerimaPulsa_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbKeySpace) Then
        Beep
        KeyAscii = 0
   End If
End Sub

Private Sub textUangBayar_Change()
    If textUangBayar.Text = "-" Then
        textTanggalBayar.Text = "Belum Bayar"
        cmbStatusBayar.ListIndex = 1
    ElseIf textUangBayar.Text = "" Then
        textTanggalBayar.Text = ""
        cmbStatusBayar.ListIndex = 2
    Else
        textTanggalBayar.Text = ""
        cmbStatusBayar.ListIndex = 0
    End If
End Sub

Private Sub textUangBayar_DblClick()
    R = SendMessageLong(CMBDataLalu6.hwnd, CB_SHOWDROPDOWN, True, 0)
End Sub

