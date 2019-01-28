VERSION 5.00
Object = "{62E1564C-1E05-44A7-A921-FD8347F324A5}#1.0#0"; "HookMenu.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Begin VB.Form FormAturPreset 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atur Preset"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   11535
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormAturPreset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   4560
      Top             =   480
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   1
      Bmp:1           =   "FormAturPreset.frx":000C
      Key:1           =   "#menuRefresh"
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
      Left            =   8760
      Top             =   1560
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
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10060
      _ExtentX        =   17754
      _ExtentY        =   3413
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
   Begin isButton3.isButton cmEdit 
      Height          =   495
      Left            =   10080
      TabIndex        =   1
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Icon            =   "FormAturPreset.frx":0434
      Style           =   1
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
      Left            =   10080
      TabIndex        =   2
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Icon            =   "FormAturPreset.frx":058E
      Style           =   1
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
   Begin isButton3.isButton cmTutup 
      Height          =   495
      Left            =   10080
      TabIndex        =   3
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Icon            =   "FormAturPreset.frx":06E8
      Style           =   1
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
   Begin isButton3.isButton cmDataAwal 
      Height          =   495
      Left            =   10080
      TabIndex        =   4
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      Style           =   2
      Caption         =   "<<"
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
   Begin isButton3.isButton cmDataSebelumnya 
      Height          =   495
      Left            =   10440
      TabIndex        =   5
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      Style           =   2
      Caption         =   "<"
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
   Begin isButton3.isButton cmDataSelanjutnya 
      Height          =   495
      Left            =   10800
      TabIndex        =   6
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      Style           =   2
      Caption         =   ">"
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
   Begin isButton3.isButton cmDataAkhir 
      Height          =   495
      Left            =   11160
      TabIndex        =   7
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      Style           =   2
      Caption         =   ">>"
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
   Begin VB.Menu menuTool 
      Caption         =   "Tool"
      Begin VB.Menu menuRefresh 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "FormAturPreset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub AturKontrol()
    Nyambungg
    With AdodcMain
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from TablePresetFilter"
        Set DataGrid1.DataSource = AdodcMain
        .Refresh
    End With
    With DataGrid1
        .Columns(0).Width = 5000
        .Columns(1).Width = 5000
        .AllowUpdate = False
    End With
    menuTool.Visible = False
    RemoveCancelMenuItem Me
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
    If AdodcMain.Recordset.RecordCount = 0 Then
        MsgBox "Tidak ada data yang akan diedit!", vbExclamation + vbOKOnly, ""
    Else
        With FormTambahPreset
            .textNamaPreset.Text = AdodcMain.Recordset.Fields(0).Value
            .textSyntaxSQL.Text = AdodcMain.Recordset.Fields(1).Value
            .Caption = "Edit Preset (Nama Tabel : TablePulsa)"
            .cmSimpan.Caption = "  &Update"
            .Show vbModal, Me
        End With
    End If
End Sub

Private Sub cmHapus_Click()
    If AdodcMain.Recordset.RecordCount = 0 Then
        MsgBox "Tidak ada data yang akan dihapus!", vbExclamation + vbOKOnly, ""
    Else
        X = MsgBox("Anda yakin ingin menghapus data ini?", vbQuestion + vbYesNo, "Hapus Data?")
        If X = vbYes Then
            With AdodcMain
                .Recordset.Delete
                .Refresh
                .Refresh
            End With
        End If
    End If
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

Private Sub menuRefresh_Click()
    AturKontrol
End Sub
