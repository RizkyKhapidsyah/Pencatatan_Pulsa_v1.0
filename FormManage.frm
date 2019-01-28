VERSION 5.00
Object = "{62E1564C-1E05-44A7-A921-FD8347F324A5}#1.0#0"; "HookMenu.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "FUSIONButtons.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormManage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   14910
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormManage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   14910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2400
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBawah 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   14
      Top             =   6795
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   15
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel13 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel14 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel15 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   13
      Top             =   6765
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   53
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   5520
      Top             =   720
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   9
      Bmp:1           =   "FormManage.frx":0442
      Key:1           =   "#menuMED"
      Bmp:2           =   "FormManage.frx":086A
      Key:2           =   "#menuMWD"
      Bmp:3           =   "FormManage.frx":0C92
      Key:3           =   "#menuHtmlDocument"
      Bmp:4           =   "FormManage.frx":10BA
      Key:4           =   "#menuBerkasLaporan"
      Bmp:5           =   "FormManage.frx":14E2
      Key:5           =   "#menuSingleFilter"
      Bmp:6           =   "FormManage.frx":190A
      Key:6           =   "#menuPreset"
      Bmp:7           =   "FormManage.frx":1D32
      Key:7           =   "#menuEksporBaris"
      Bmp:8           =   "FormManage.frx":215A
      Key:8           =   "#menuCBP"
      Bmp:9           =   "FormManage.frx":2582
      Key:9           =   "#menuSBP"
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
      Left            =   0
      Top             =   6720
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
   Begin XPEngine.XPControl XPEngine 
      Left            =   120
      Top             =   6720
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin isButton3.isButton cmEdit 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormManage.frx":29AA
      Style           =   6
      Caption         =   "&Edit"
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
   Begin isButton3.isButton cmCari 
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormManage.frx":2B04
      Style           =   6
      Caption         =   "&Cari"
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
   Begin isButton3.isButton cmSorot 
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormManage.frx":3096
      Style           =   6
      Caption         =   "&Sorot"
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
   Begin isButton3.isButton cmFilter 
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormManage.frx":31F0
      Style           =   6
      Caption         =   "&Filter"
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
   Begin isButton3.isButton cmHapus 
      Height          =   495
      Left            =   6720
      TabIndex        =   5
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormManage.frx":334A
      Style           =   6
      Caption         =   "&Hapus"
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
   Begin isButton3.isButton cmRefresh 
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormManage.frx":34A4
      Style           =   6
      Caption         =   "&Refresh"
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
   Begin KewlButtonz.KewlButtons cmDataAwal 
      Height          =   495
      Left            =   12840
      TabIndex        =   7
      Top             =   6960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   7
      TX              =   "<<"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormManage.frx":35FE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons cmDataSebelumnya 
      Height          =   495
      Left            =   13320
      TabIndex        =   8
      Top             =   6960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   7
      TX              =   "<"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormManage.frx":361A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons cmDataSelanjutnya 
      Height          =   495
      Left            =   13800
      TabIndex        =   9
      Top             =   6960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   7
      TX              =   ">"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormManage.frx":3636
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons cmDataAkhir 
      Height          =   495
      Left            =   14280
      TabIndex        =   10
      Top             =   6960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   7
      TX              =   ">>"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormManage.frx":3652
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin isButton3.isButton cmEskpor 
      Height          =   495
      Left            =   8040
      TabIndex        =   11
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormManage.frx":366E
      Style           =   6
      Caption         =   "   E&kspor"
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
   Begin isButton3.isButton cmSQL 
      Height          =   495
      Left            =   9360
      TabIndex        =   12
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormManage.frx":3988
      Style           =   6
      Caption         =   "    SQL Exec"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   9763
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   21
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin isButton3.isButton cmTutup 
      Height          =   495
      Left            =   10680
      TabIndex        =   15
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormManage.frx":3AE2
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
   Begin VB.Menu menuFilter 
      Caption         =   "Filter"
      Begin VB.Menu menuSingleFilter 
         Caption         =   "Single Filter"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu menuPreset 
         Caption         =   "Preset"
      End
   End
   Begin VB.Menu menuEkspor 
      Caption         =   "Ekspor"
      Begin VB.Menu menuMED 
         Caption         =   "Microsoft Excel Document"
      End
      Begin VB.Menu menuEksporBaris 
         Caption         =   "Ekspor Baris Ini"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu menuCBP 
         Caption         =   "Cetak Bon Pembayaran"
      End
   End
End
Attribute VB_Name = "FormManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    On Error GoTo HancurkanError
    Nyambungg
    With AdodcMain
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From TablePulsa order by Tahun asc, Bulan, Tanggal"
        Set DataGrid1.DataSource = AdodcMain
        .Refresh
    End With
    AturUkuranDatagrid
    MasukkanDataKeStatus
    If AdodcMain.Recordset.RecordCount = 0 Then Else AdodcMain.Recordset.MoveLast
    RemoveCancelMenuItem Me
    XPEngine.StartEngine
        Me.Picture = LoadPicture(App.Path & "\image\bannerManage.jpg")
        menuEkspor.Visible = False
        menuFilter.Visible = False
        If FormPengaturan.CekStatusBawahFormManage.Value = Checked Then
            Me.StatusBawah.Visible = True
            Me.Height = 8480
        ElseIf FormPengaturan.CekStatusBawahFormManage.Value = Unchecked Then
            Me.StatusBawah.Visible = False
            Me.Height = 7980
        End If
    AturThema
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
Sub AturUkuranDatagrid()
If AdodcMain.Recordset.Fields.Count = 15 Then
    With DataGrid1
        .Columns(0).Width = 555.0236
        .Columns(1).Width = 629.8583
        .Columns(2).Width = 494.9292
        .Columns(3).Width = 524.9764
        .Columns(4).Width = 1154.835
        .Columns(5).Width = 1400.835
        .Columns(6).Width = 780.0945
        .Columns(7).Width = 1110.047
        .Columns(8).Width = 1440
        .Columns(9).Width = 1019.906
        .Columns(10).Width = 870.2363
        .Columns(11).Width = 1100.8819
        .Columns(12).Width = 1700.906
        .Columns(13).Width = 464.8819
        .Columns(14).Width = 1019.906
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Alignment = dbgCenter
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Alignment = dbgCenter
        .Columns(7).Alignment = dbgCenter
        .Columns(8).Alignment = dbgCenter
        .Columns(9).Alignment = dbgCenter
        .Columns(10).Alignment = dbgCenter
        .Columns(11).Alignment = dbgCenter
        .Columns(12).Alignment = dbgCenter
        .Columns(13).Alignment = dbgCenter
        .Columns(14).Alignment = dbgCenter
            
        If FormPengaturan.cekKunciTabel.Value = Checked Then
            .AllowUpdate = False
        ElseIf FormPengaturan.cekKunciTabel.Value = Unchecked Then
            .AllowUpdate = True
        End If
    End With
Else
    With DataGrid1
        .Columns(0).Width = 1000
    End With
End If
End Sub
Sub MasukkanDataKeStatus()
On Error Resume Next
With StatusBawah
    .Panels.Item(1).Width = DataGrid1.Columns(0).Width
    .Panels.Item(2).Width = DataGrid1.Columns(1).Width
    .Panels.Item(3).Width = DataGrid1.Columns(2).Width
    .Panels.Item(4).Width = DataGrid1.Columns(3).Width
    .Panels.Item(5).Width = DataGrid1.Columns(4).Width
    .Panels.Item(6).Width = DataGrid1.Columns(5).Width
    .Panels.Item(7).Width = DataGrid1.Columns(6).Width
    .Panels.Item(8).Width = DataGrid1.Columns(7).Width
    .Panels.Item(9).Width = DataGrid1.Columns(8).Width
    .Panels.Item(10).Width = DataGrid1.Columns(9).Width
    .Panels.Item(11).Width = DataGrid1.Columns(10).Width
    .Panels.Item(12).Width = DataGrid1.Columns(11).Width
    .Panels.Item(13).Width = DataGrid1.Columns(12).Width
    .Panels.Item(14).Width = DataGrid1.Columns(13).Width
    .Panels.Item(15).Width = DataGrid1.Columns(14).Width
    
    .Panels.Item(1).Alignment = sbrLeft
    .Panels.Item(2).Alignment = sbrCenter
    .Panels.Item(3).Alignment = sbrCenter
    .Panels.Item(4).Alignment = sbrCenter
    .Panels.Item(5).Alignment = sbrCenter
    .Panels.Item(6).Alignment = sbrCenter
    .Panels.Item(7).Alignment = sbrCenter
    .Panels.Item(8).Alignment = sbrCenter
    .Panels.Item(9).Alignment = sbrCenter
    .Panels.Item(10).Alignment = sbrCenter
    .Panels.Item(11).Alignment = sbrCenter
    .Panels.Item(12).Alignment = sbrCenter
    .Panels.Item(13).Alignment = sbrCenter
    .Panels.Item(14).Alignment = sbrCenter
    .Panels.Item(15).Alignment = sbrCenter
    
    .Panels.Item(1).Text = AdodcMain.Recordset.Fields(0).Value
    .Panels.Item(2).Text = AdodcMain.Recordset.Fields(1).Value
    .Panels.Item(3).Text = AdodcMain.Recordset.Fields(2).Value
    .Panels.Item(4).Text = AdodcMain.Recordset.Fields(3).Value
    .Panels.Item(5).Text = AdodcMain.Recordset.Fields(4).Value
    .Panels.Item(6).Text = AdodcMain.Recordset.Fields(5).Value
    .Panels.Item(7).Text = AdodcMain.Recordset.Fields(6).Value
    .Panels.Item(8).Text = AdodcMain.Recordset.Fields(7).Value
    .Panels.Item(9).Text = AdodcMain.Recordset.Fields(8).Value
    .Panels.Item(10).Text = AdodcMain.Recordset.Fields(9).Value
    .Panels.Item(11).Text = AdodcMain.Recordset.Fields(10).Value
    .Panels.Item(12).Text = AdodcMain.Recordset.Fields(11).Value
    .Panels.Item(13).Text = AdodcMain.Recordset.Fields(12).Value
    .Panels.Item(14).Text = AdodcMain.Recordset.Fields(13).Value
    .Panels.Item(15).Text = AdodcMain.Recordset.Fields(14).Value

    .Panels.Item(1).ToolTipText = AdodcMain.Recordset.Fields(0).Name & " = " & .Panels.Item(1).Text
    .Panels.Item(2).ToolTipText = AdodcMain.Recordset.Fields(1).Name & " = " & .Panels.Item(2).Text
    .Panels.Item(3).ToolTipText = AdodcMain.Recordset.Fields(2).Name & " = " & .Panels.Item(3).Text
    .Panels.Item(4).ToolTipText = AdodcMain.Recordset.Fields(3).Name & " = " & .Panels.Item(4).Text
    .Panels.Item(5).ToolTipText = AdodcMain.Recordset.Fields(4).Name & " = " & .Panels.Item(5).Text
    .Panels.Item(6).ToolTipText = AdodcMain.Recordset.Fields(5).Name & " = " & .Panels.Item(6).Text
    .Panels.Item(7).ToolTipText = AdodcMain.Recordset.Fields(6).Name & " = " & .Panels.Item(7).Text
    .Panels.Item(8).ToolTipText = AdodcMain.Recordset.Fields(7).Name & " = " & .Panels.Item(8).Text
    .Panels.Item(9).ToolTipText = AdodcMain.Recordset.Fields(8).Name & " = " & .Panels.Item(9).Text
    .Panels.Item(10).ToolTipText = AdodcMain.Recordset.Fields(9).Name & " = " & .Panels.Item(10).Text
    .Panels.Item(11).ToolTipText = AdodcMain.Recordset.Fields(10).Name & " = " & .Panels.Item(11).Text
    .Panels.Item(12).ToolTipText = AdodcMain.Recordset.Fields(11).Name & " = " & .Panels.Item(12).Text
    .Panels.Item(13).ToolTipText = AdodcMain.Recordset.Fields(12).Name & " = " & .Panels.Item(13).Text
    .Panels.Item(14).ToolTipText = AdodcMain.Recordset.Fields(13).Name & " = " & .Panels.Item(14).Text
    .Panels.Item(15).ToolTipText = AdodcMain.Recordset.Fields(14).Name & " = " & .Panels.Item(15).Text

End With
End Sub

Private Sub cmCari_Click()
    If AdodcMain.Recordset.RecordCount = 0 Then
        MsgBox "Tidak ada data yang akan dicari!", vbExclamation + vbOKOnly, ""
    Else
        If AdodcMain.Recordset.Fields.Count = 15 Then
            FormCari.Show vbModal, Me
        Else
            MsgBox "Silankan klik Refresh, kemudian cari data yang akan dicari!", vbExclamation + vbOKOnly, "Mohon Refresh Data"
            cmRefresh.SetFocus
        End If
    End If
End Sub

Private Sub cmDataAkhir_Click()
    AdodcMain.Recordset.MoveLast
    MasukkanDataKeStatus
End Sub

Private Sub cmDataAwal_Click()
    AdodcMain.Recordset.MoveFirst
    MasukkanDataKeStatus
End Sub

Private Sub cmDataSebelumnya_Click()
    AdodcMain.Recordset.MovePrevious
    If AdodcMain.Recordset.BOF = True Then AdodcMain.Recordset.MoveLast
    MasukkanDataKeStatus
End Sub

Private Sub cmDataSelanjutnya_Click()
    AdodcMain.Recordset.MoveNext
    If AdodcMain.Recordset.EOF = True Then AdodcMain.Recordset.MoveFirst
    MasukkanDataKeStatus
End Sub

Private Sub cmEdit_Click()
If AdodcMain.Recordset.RecordCount = 0 Then
    MsgBox "Tidak ada data yang akan diedit!", vbExclamation + vbOKOnly, ""
Else
    If AdodcMain.Recordset.Fields.Count = 15 Then
        With FormBaru
            .Caption = "Edit Data"
            .cmbHari.Text = AdodcMain.Recordset.Fields(0).Value
            .cmbTanggal.Text = AdodcMain.Recordset.Fields(1).Value
            .cmbBulan.Text = AdodcMain.Recordset.Fields(2).Value
            .cmbTahun.Text = AdodcMain.Recordset.Fields(3).Value
            .textPenerimaPulsa.Text = AdodcMain.Recordset.Fields(4).Value
            .TextNamaPenerima.Text = AdodcMain.Recordset.Fields(5).Value
            .cmbProvider.Text = AdodcMain.Recordset.Fields(6).Value
            .cmbJenisProvider.Text = AdodcMain.Recordset.Fields(7).Value
            .textJumlahPulsaDibeli.Text = AdodcMain.Recordset.Fields(8).Value
            .textHargaServer.Text = AdodcMain.Recordset.Fields(9).Value
            .TextHargaJual.Text = AdodcMain.Recordset.Fields(10).Value
            .textUangBayar.Text = AdodcMain.Recordset.Fields(11).Value
            .textTanggalBayar.Text = AdodcMain.Recordset.Fields(12).Value
            .textLaba.Text = AdodcMain.Recordset.Fields(13).Value
            .cmbStatusBayar.Text = AdodcMain.Recordset.Fields(14).Value
            .cmSimpan.Caption = "&Update"
            .Show vbModal, Me
        End With
    Else
        MsgBox "Silankan klik Refresh, kemudian cari data yang akan diedit!", vbExclamation + vbOKOnly, "Mohon Refresh Data"
        cmRefresh.SetFocus
    End If
End If
End Sub

Private Sub cmEskpor_Click()
    PopupMenu menuEkspor
End Sub

Private Sub cmFilter_Click()
    PopupMenu menuFilter
End Sub

Private Sub cmHapus_Click()
    If AdodcMain.Recordset.RecordCount = 0 Then
        MsgBox "Tidak ada data yang akan dihapus!", vbExclamation + vbOKOnly, ""
    Else
        X = MsgBox("Berikut keterangan data yang akan dihapus : " & vbCrLf & vbCrLf & _
                    "===============================================" & vbCrLf & _
                    "Tanggal : " & AdodcMain.Recordset.Fields(0).Value & ", " & AdodcMain.Recordset.Fields(1).Value & " - " & AdodcMain.Recordset.Fields(2).Value & " - " & AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(4).Name & " : " & AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(5).Name & " : " & AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(6).Name & " : " & AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(7).Name & " : " & AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(8).Name & " : " & AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(9).Name & " : " & AdodcMain.Recordset.Fields(9).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(10).Name & " : " & AdodcMain.Recordset.Fields(10).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(11).Name & " : " & AdodcMain.Recordset.Fields(11).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(12).Name & " : " & AdodcMain.Recordset.Fields(12).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(13).Name & " : " & AdodcMain.Recordset.Fields(13).Value & vbCrLf & _
                    AdodcMain.Recordset.Fields(14).Name & " : " & AdodcMain.Recordset.Fields(14).Value & vbCrLf & _
                    "===============================================" & vbCrLf & vbCrLf & _
                    "Yakin untuk menghapus data ini?", vbQuestion + vbYesNo, "Hapus?")
        If X = vbYes Then
            With AdodcMain
                .Recordset.Delete
                .Refresh
                .Refresh
            End With
            FormManage.AturUkuranDatagrid
            FormUtama.cmRefresh_Click
            cmRefresh_Click
        End If
    End If
End Sub

Public Sub cmRefresh_Click()
    AturKontrol
    cmEdit.Enabled = True
    cmCari.Enabled = True
    cmSorot.Enabled = True
    cmHapus.Enabled = True
    cmEskpor.Enabled = True
End Sub

Private Sub cmSorot_Click()
    If AdodcMain.Recordset.RecordCount = 0 Then
        MsgBox "Tidak ada data yang akan disorot!", vbExclamation + vbOKOnly, ""
    Else
        If AdodcMain.Recordset.Fields.Count = 15 Then
            FormSorot.Show vbModal, Me
        Else
            MsgBox "Silankan klik Refresh, kemudian cari data yang akan disorot!", vbExclamation + vbOKOnly, "Mohon Refresh Data"
            cmRefresh.SetFocus
        End If
    End If
End Sub

Private Sub cmSQL_Click()
    FormExecuteSQL.Show vbModal, Me
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub DataGrid1_Click()
    MasukkanDataKeStatus
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

Private Sub menuCBP_Click()
    If AdodcMain.Recordset.Fields(14).Value = "Hutang" Or AdodcMain.Recordset.Fields(14).Value = "Menunggu" Then
        MsgBox "Maaf, bon tidak dapat dicetak karena pada list, yang bersangkutan status bayar masih 'Hutang' atau 'Menunggu' pembayaran!", vbExclamation + vbOKOnly, "Bon Tidak Dapat Dicetak"
    Else
        FormCetakBonPembayaran.Show vbModal, Me
    End If
End Sub

Private Sub menuEksporBaris_Click()
    On Error GoTo ErrorHandler
    CommonDialog1.DialogTitle = "Ekspor Baris"
    CommonDialog1.FileName = AdodcMain.Recordset.Fields(5).Value & " (" & AdodcMain.Recordset.Fields(4).Value & ") - " & AdodcMain.Recordset.Fields(0).Value & ", " & AdodcMain.Recordset.Fields(1).Value & "-" & AdodcMain.Recordset.Fields(2).Value & "-" & AdodcMain.Recordset.Fields(3).Value
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
    Print #iFile, AdodcMain.Recordset.Fields(0).Name & " : " & AdodcMain.Recordset.Fields(0).Value
    Print #iFile, AdodcMain.Recordset.Fields(1).Name & " : " & AdodcMain.Recordset.Fields(1).Value
    Print #iFile, AdodcMain.Recordset.Fields(2).Name & " : " & AdodcMain.Recordset.Fields(2).Value
    Print #iFile, AdodcMain.Recordset.Fields(3).Name & " : " & AdodcMain.Recordset.Fields(3).Value
    Print #iFile, AdodcMain.Recordset.Fields(4).Name & " : " & AdodcMain.Recordset.Fields(4).Value
    Print #iFile, AdodcMain.Recordset.Fields(5).Name & " : " & AdodcMain.Recordset.Fields(5).Value
    Print #iFile, AdodcMain.Recordset.Fields(6).Name & " : " & AdodcMain.Recordset.Fields(6).Value
    Print #iFile, AdodcMain.Recordset.Fields(7).Name & " : " & AdodcMain.Recordset.Fields(7).Value
    Print #iFile, AdodcMain.Recordset.Fields(8).Name & " : " & AdodcMain.Recordset.Fields(8).Value
    Print #iFile, AdodcMain.Recordset.Fields(9).Name & " : " & AdodcMain.Recordset.Fields(9).Value
    Print #iFile, AdodcMain.Recordset.Fields(10).Name & " : " & AdodcMain.Recordset.Fields(10).Value
    Print #iFile, AdodcMain.Recordset.Fields(11).Name & " : " & AdodcMain.Recordset.Fields(11).Value
    Print #iFile, AdodcMain.Recordset.Fields(12).Name & " : " & AdodcMain.Recordset.Fields(12).Value
    Print #iFile, AdodcMain.Recordset.Fields(13).Name & " : " & AdodcMain.Recordset.Fields(13).Value
    Print #iFile, AdodcMain.Recordset.Fields(14).Name & " : " & AdodcMain.Recordset.Fields(14).Value
    SaveFileFromTB = True
ErrorHandler:
    Close #iFile
End Sub

Private Sub menuMED_Click()
    FormLoading.Show vbModal, Me
End Sub




Private Sub menuPreset_Click()
    If AdodcMain.Recordset.RecordCount = 0 Then
        MsgBox "Tidak ada data yang akan difilter!", vbExclamation + vbOKOnly, ""
    Else
        cmRefresh_Click
        FormPreset.Show vbModal, Me
    End If
End Sub

Private Sub menuSingleFilter_Click()
    If AdodcMain.Recordset.RecordCount = 0 Then
        MsgBox "Tidak ada data yang akan difilter!", vbExclamation + vbOKOnly, ""
    Else
        cmRefresh_Click
        FormFilter.Show vbModal, Me
    End If
End Sub
