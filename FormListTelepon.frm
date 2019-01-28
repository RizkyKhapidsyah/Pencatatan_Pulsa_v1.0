VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "FUSIONButtons.ocx"
Begin VB.Form FormListTelepon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List Telepon"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7185
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormListTelepon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdodcMain 
      Height          =   330
      Left            =   -120
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox textData 
      Height          =   480
      Left            =   1125
      TabIndex        =   7
      Top             =   3735
      Width           =   3615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   21
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9.75
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
   Begin XPEngine.XPControl XPEngine 
      Left            =   240
      Top             =   2520
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin isButton3.isButton cmEdit 
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormListTelepon.frx":57E2
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
      Left            =   5880
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormListTelepon.frx":593C
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
   Begin isButton3.isButton cmHapus 
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormListTelepon.frx":5ECE
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
      Left            =   5880
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormListTelepon.frx":6028
      Style           =   6
      Caption         =   "  &Refresh"
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
   Begin isButton3.isButton cmTambah 
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormListTelepon.frx":6182
      Style           =   6
      Caption         =   "  &Tambah"
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
      Left            =   5880
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormListTelepon.frx":62DC
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
   Begin KewlButtonz.KewlButtons cmDataAwal 
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   8
      TX              =   "<<"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "FormListTelepon.frx":6436
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons cmDataSebelumnya 
      Height          =   495
      Left            =   612
      TabIndex        =   9
      Top             =   3720
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   8
      TX              =   "<"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "FormListTelepon.frx":6452
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons cmDataSelanjutnya 
      Height          =   495
      Left            =   4755
      TabIndex        =   10
      Top             =   3720
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   8
      TX              =   ">"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "FormListTelepon.frx":646E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons cmDataAkhir 
      Height          =   495
      Left            =   5250
      TabIndex        =   11
      Top             =   3720
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   8
      TX              =   ">>"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "FormListTelepon.frx":648A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin isButton3.isButton cmMasukan 
      Height          =   495
      Left            =   5880
      TabIndex        =   12
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormListTelepon.frx":64A6
      Style           =   6
      Caption         =   "    &Masukkan"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FormListTelepon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    Nyambungg
    With AdodcMain
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From TableListTelepon Order by Nama Asc;"
        Set DataGrid1.DataSource = AdodcMain
        .Refresh
        .Recordset.MoveLast
    End With
    AturUkuranDatagrid
    RemoveCancelMenuItem Me
    DataKeTextBox
    XPEngine.StartEngine
    If FormPengaturan.cekKunciTabel.Value = Checked Then
        DataGrid1.AllowUpdate = False
    ElseIf FormPengaturan.cekKunciTabel.Value = Unchecked Then
        DataGrid1.AllowUpdate = True
    End If
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
Sub AturUkuranDatagrid()
    With DataGrid1
        .Columns(0).Width = 3800
        .Columns(1).Width = 1200
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Alignment = dbgCenter
    End With
End Sub
Sub DataKeTextBox()
    With textData
        .Text = "(Data Ke : '" & AdodcMain.Recordset.AbsolutePosition & "' Dari '" & AdodcMain.Recordset.RecordCount & "' data) - (" & AdodcMain.Recordset.Fields(0).Value & " (" & AdodcMain.Recordset.Fields(1).Value & "))"
        .Alignment = 2
        .Locked = True
        .ToolTipText = .Text
    End With
End Sub

Private Sub cmCari_Click()
    FormCariListTelepon.Show vbModal, Me
End Sub

Private Sub cmDataAkhir_Click()
    AdodcMain.Recordset.MoveLast
    DataKeTextBox
End Sub

Private Sub cmDataAwal_Click()
    AdodcMain.Recordset.MoveFirst
    DataKeTextBox
End Sub

Private Sub cmDataSebelumnya_Click()
    AdodcMain.Recordset.MovePrevious
    If AdodcMain.Recordset.BOF = True Then AdodcMain.Recordset.MoveLast
    DataKeTextBox
End Sub

Private Sub cmDataSelanjutnya_Click()
    AdodcMain.Recordset.MoveNext
    If AdodcMain.Recordset.EOF = True Then AdodcMain.Recordset.MoveFirst
    DataKeTextBox
End Sub

Private Sub cmEdit_Click()
    If AdodcMain.Recordset.RecordCount = 0 Then
        MsgBox "Tidak ada data yang akan diedit!", vbExclamation + vbOKOnly, ""
    Else
        With FormTambahEditListTelepon
            .Caption = "Edit"
            .cmSimpan.Caption = "&Update"
            .textNama.Text = AdodcMain.Recordset.Fields(0).Value
            .textNomorTelepon.Text = AdodcMain.Recordset.Fields(1).Value
            .Show vbModal, Me
        End With
    End If
End Sub

Private Sub cmHapus_Click()
    If AdodcMain.Recordset.RecordCount = 0 Then
        MsgBox "Tidak ada data yang akan dihapus!", vbExclamation + vbOKOnly, ""
    Else
        X = MsgBox("Anda yakin ingin menghapus data ini?", vbQuestion + vbYesNo, "Hapus?")
        If X = vbYes Then
            With AdodcMain
                .Recordset.Delete
                .Refresh
                .Refresh
                .Recordset.MoveLast
            End With
            AturUkuranDatagrid
        End If
    End If
End Sub

Private Sub cmMasukan_Click()
On Error GoTo HancurkanError
    With FormBaru
        .TextNamaPenerima.Text = AdodcMain.Recordset.Fields(0).Value
        .textPenerimaPulsa.Text = AdodcMain.Recordset.Fields(1).Value
    End With
    Unload Me
    FormBaru.textJumlahPulsaDibeli.SetFocus
Exit Sub
HancurkanError:
    PusatError
End Sub

Private Sub cmRefresh_Click()
    AturKontrol
End Sub

Private Sub cmTambah_Click()
    With FormTambahEditListTelepon
        .Caption = "Tambah"
        .cmSimpan.Caption = "&Simpan"
        .Show vbModal, Me
    End With
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub
