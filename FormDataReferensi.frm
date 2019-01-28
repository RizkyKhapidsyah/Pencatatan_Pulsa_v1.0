VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "FUSIONButtons.ocx"
Begin VB.Form FormDataReferensi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Referensi"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5385
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormDataReferensi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin KewlButtonz.KewlButtons cmDataAwal 
      Height          =   300
      Left            =   3840
      TabIndex        =   6
      Top             =   3600
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
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
      MICON           =   "FormDataReferensi.frx":014A
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
      Height          =   300
      Left            =   4200
      TabIndex        =   7
      Top             =   3600
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
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
      MICON           =   "FormDataReferensi.frx":0166
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
      Height          =   300
      Left            =   4560
      TabIndex        =   8
      Top             =   3600
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
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
      MICON           =   "FormDataReferensi.frx":0182
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
      Height          =   300
      Left            =   4920
      TabIndex        =   9
      Top             =   3600
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
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
      MICON           =   "FormDataReferensi.frx":019E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin XPEngine.XPControl XPEngine 
      Left            =   5280
      Top             =   480
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin VB.CheckBox cekKunciTabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Kunci Tabel"
      Height          =   270
      Left            =   4080
      TabIndex        =   11
      Top             =   360
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc AdodcMain 
      Height          =   330
      Left            =   7920
      Top             =   4560
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5741
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
   Begin VB.ComboBox cmbDataReferensi 
      Height          =   390
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin isButton3.isButton cmEdit 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormDataReferensi.frx":01BA
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
   Begin isButton3.isButton cmHapus 
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormDataReferensi.frx":0314
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
      Left            =   1440
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormDataReferensi.frx":046E
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
   Begin isButton3.isButton cmTutup 
      Height          =   495
      Left            =   4080
      TabIndex        =   10
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormDataReferensi.frx":05C8
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pilih Referensi :"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "FormDataReferensi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cekKunciTabel_Click()
    If cekKunciTabel.Value = Checked Then
        DataGrid1.AllowUpdate = False
    ElseIf cekKunciTabel.Value = Unchecked Then
        DataGrid1.AllowUpdate = True
    End If
End Sub

Private Sub cmbDataReferensi_Click()
    Select Case cmbDataReferensi.ListIndex
    Case Is = 0
        With AdodcMain
            .ConnectionString = CN.ConnectionString
            .RecordSource = "Select * from TableReferensiHargaJual order by Harga_Jual asc;"
            .Refresh
        End With
    Case Is = 1
        With AdodcMain
            .ConnectionString = CN.ConnectionString
            .RecordSource = "Select * from TableReferensiHargaServer order by Harga_Server asc;"
            .Refresh
        End With
    Case Is = 2
        With AdodcMain
            .ConnectionString = CN.ConnectionString
            .RecordSource = "Select * from TableReferensiJumlahPulsaDibeli order by Jumlah_Pulsa_Dibeli asc;"
            .Refresh
        End With
    End Select
        Set DataGrid1.DataSource = AdodcMain
        DataGrid1.Columns.Item(0).Width = 3000
End Sub

Private Sub cmDataAkhir_Click()
    AdodcMain.Recordset.MoveLast
End Sub

Private Sub cmDataAwal_Click()
    AdodcMain.Recordset.MoveFirst
End Sub

Private Sub cmDataSebelumnya_Click()
    AdodcMain.Recordset.MovePrevious
    If AdodcMain.Recordset.BOF = True Then AdodcMain.Recordset.MoveLast
End Sub

Private Sub cmDataSelanjutnya_Click()
    AdodcMain.Recordset.MoveNext
    If AdodcMain.Recordset.EOF = True Then AdodcMain.Recordset.MoveFirst
End Sub


Private Sub cmEdit_Click()
    With FormEditDataReferensi
        Select Case cmbDataReferensi.ListIndex
            Case Is = 0
                .Caption = "Edit Ref. Harga Jual"
                .LabelEdit.Caption = "Edit Nilai :"
                .textEdit.Text = AdodcMain.Recordset.Fields(0).Value
            Case Is = 1
                .Caption = "Edit Ref. Harga Server"
                .LabelEdit.Caption = "Edit Nilai :"
                .textEdit.Text = AdodcMain.Recordset.Fields(0).Value
            Case Is = 2
                .Caption = "Edit Ref. Jumlah Dibeli"
                .LabelEdit.Caption = "Edit Nilai :"
                .textEdit.Text = AdodcMain.Recordset.Fields(0).Value
            End Select
            .Show vbModal, Me
    End With
End Sub

Private Sub cmHapus_Click()
    If AdodcMain.Recordset.RecordCount = 0 Then
        MsgBox "Tidak ada data yang akan dihapus!", vbExclamation + vbOKOnly, ""
    Else
        X = MsgBox("Yakin untuk menghapus data ini?", vbQuestion + vbYesNo, "Hapus?")
        If X = vbYes Then
            With AdodcMain
                .Recordset.Delete
                .Refresh
                .Refresh
            End With
            Form_Load
        End If
    End If
End Sub

Private Sub cmRefresh_Click()
    Form_Load
End Sub

Private Sub cmTutup_Click()
    SaveSetting App.Title, "RSBR\DataRef", Me.cekKunciTabel.Name, Me.cekKunciTabel.Value
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    MsgBox "Jumlah Data : " & AdodcMain.Recordset.RecordCount, vbInformation + vbOKOnly, "Jumlah Data"
End Sub

Private Sub Form_Load()
    Nyambungg
    XPEngine.StartEngine
    AturThema
    RemoveCancelMenuItem Me
    R = SendMessageLong(cmbDataReferensi.hwnd, CB_SETDROPPEDWIDTH, 200, 0)
    With cmbDataReferensi
        .Clear
        .AddItem "Referensi Harga Jual", 0
        .AddItem "Referensi Harga Server", 1
        .AddItem "Referensi Jumlah Pulsa Dibeli", 2
        .ListIndex = 0
    End With
    DataGrid1.Columns.Item(0).Width = 3000
        Me.cekKunciTabel.Value = GetSetting(App.Title, "RSBR\DataRef", Me.cekKunciTabel.Name, Me.cekKunciTabel.Value)
    If cekKunciTabel.Value = Checked Then
        DataGrid1.AllowUpdate = False
    ElseIf cekKunciTabel.Value = Unchecked Then
        DataGrid1.AllowUpdate = True
    End If
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
