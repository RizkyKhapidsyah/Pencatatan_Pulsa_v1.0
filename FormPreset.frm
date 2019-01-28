VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "FUSIONButtons.ocx"
Begin VB.Form FormPreset 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preset Filter"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormPreset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox CekTutupForm 
      Caption         =   "Tutup jendela ini setelah data difilter"
      Height          =   270
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Value           =   1  'Checked
      Width           =   3615
   End
   Begin MSAdodcLib.Adodc AdodcMain 
      Height          =   330
      Left            =   2040
      Top             =   120
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
   Begin VB.TextBox TextSQL 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "FormPreset.frx":000C
      Top             =   1920
      Width           =   6735
   End
   Begin XPEngine.XPControl XPEngine 
      Left            =   120
      Top             =   3240
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin VB.ComboBox cmbPreset 
      Height          =   390
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   6255
   End
   Begin isButton3.isButton cmFilter 
      Height          =   495
      Left            =   5790
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Icon            =   "FormPreset.frx":0012
      Style           =   6
      Caption         =   "  &Filter"
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
      Left            =   4560
      TabIndex        =   3
      Top             =   1320
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   873
      Icon            =   "FormPreset.frx":016C
      Style           =   6
      Caption         =   "  &Tutup"
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
   Begin isButton3.isButton cmTambahPreset 
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   480
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   661
      Style           =   6
      Caption         =   "+"
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
   Begin KewlButtonz.KewlButtons cmTampilkanSQL 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   7
      TX              =   "Tampilkan SQL"
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
      MICON           =   "FormPreset.frx":02C6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin isButton3.isButton cmAturPreset 
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   3120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      Icon            =   "FormPreset.frx":02E2
      Style           =   6
      Caption         =   "&Atur Preset"
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
      Caption         =   "Pilih Preset :"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "FormPreset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub AturKontrol()
    Nyambungg
    With AdodcMain
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from TablePresetFilter"
        .Refresh
    End With
    XPEngine.StartEngine
    With TextSQL
        .Locked = True
        .Text = ""
    End With
    cmTambahPreset.ToolTipText = "Tambah Preset"
    IsiCMBPreset
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
Sub IsiCMBPreset()
    With Me
        .cmbPreset.Clear
        Do Until AdodcMain.Recordset.EOF
            .cmbPreset.AddItem AdodcMain.Recordset.Fields(0).Value, 0
            AdodcMain.Recordset.MoveNext
        Loop
        .cmbPreset.ListIndex = 0
        AdodcMain.Refresh
    End With
End Sub

Private Sub cmAturPreset_Click()
    FormAturPreset.Show vbModal, Me
End Sub

Private Sub cmbPreset_Click()
Kalimat = "Nama_Preset = '" & cmbPreset.Text & "'"
AdodcMain.Refresh
    With AdodcMain.Recordset
        .Find Kalimat
        If Not .EOF Then
            TextSQL.Text = .Fields(1).Value
        End If
    End With
End Sub

Private Sub cmFilter_Click()
On Error GoTo HancurkanError
    With FormManage
        .AturKontrol
        .AdodcMain.Refresh
        .AdodcMain.RecordSource = Me.TextSQL.Text
        Set .DataGrid1.DataSource = .AdodcMain
        .AdodcMain.Refresh
        If FormManage.DataGrid1.Columns.Count = 15 Then
            .cmEdit.Enabled = True
            .cmCari.Enabled = True
            .cmSorot.Enabled = True
            .cmHapus.Enabled = True
        Else
            .cmEdit.Enabled = False
            .cmCari.Enabled = False
            .cmSorot.Enabled = False
            .cmHapus.Enabled = False
        End If
        .AturUkuranDatagrid
    End With
    If CekTutupForm.Value = Checked Then Unload Me
Exit Sub
HancurkanError:
    If Err.Number = "-2147217900" Then
        MsgBox "Syntax SQL salah!" & vbCrLf & _
                "Mohon untuk diperbaiki kembali!", vbCritical + vbOKOnly, "SQL Error"
    Else
        PusatError
    End If
    FormManage.cmRefresh_Click
    Me.AdodcMain.Refresh
    Me.AdodcMain.Refresh
End Sub

Private Sub cmTambahPreset_Click()
    With FormTambahPreset
        .Caption = "Tambah Preset (Nama Tabel : TablePulsa)"
        .Show vbModal, Me
    End With
End Sub

Private Sub cmTampilkanSQL_Click()
    Select Case cmTampilkanSQL.Caption
        Case "Tampilkan SQL"
            Me.Height = 4140
            cmTampilkanSQL.Caption = "Sembunyikan SQL"
            TextSQL.SetFocus
        Case "Sembunyikan SQL"
            Me.Height = 2325
            cmTampilkanSQL.Caption = "Tampilkan SQL"
            cmbPreset.SetFocus
    End Select
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

