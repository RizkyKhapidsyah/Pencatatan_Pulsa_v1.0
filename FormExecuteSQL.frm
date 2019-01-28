VERSION 5.00
Object = "{62E1564C-1E05-44A7-A921-FD8347F324A5}#1.0#0"; "HookMenu.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Begin VB.Form FormExecuteSQL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Execute SQL (Nama Tabel : TablePulsa)"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormExecuteSQL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   2520
      Top             =   120
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   2
      Bmp:1           =   "FormExecuteSQL.frx":014A
      Key:1           =   "#menuReset"
      Bmp:2           =   "FormExecuteSQL.frx":0572
      Key:2           =   "#menuHapusText"
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
      Left            =   1800
      Top             =   1680
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
   Begin VB.TextBox textSQL 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "FormExecuteSQL.frx":099A
      Top             =   480
      Width           =   5655
   End
   Begin XPEngine.XPControl XPEngine 
      Left            =   360
      Top             =   4200
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin isButton3.isButton cmHistorySQL 
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormExecuteSQL.frx":09A0
      Style           =   6
      Caption         =   "   &History"
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
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormExecuteSQL.frx":0CBA
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
   Begin isButton3.isButton cmExecute 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormExecuteSQL.frx":0E14
      Style           =   6
      Caption         =   "   &Execute"
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
      Caption         =   "Masukkan Syntax SQL :"
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1500
   End
   Begin VB.Menu menuTool 
      Caption         =   "Tool"
      Begin VB.Menu menuHapusText 
         Caption         =   "Hapus Text"
      End
   End
End
Attribute VB_Name = "FormExecuteSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    Nyambungg
    With AdodcMain
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from TableHistorySQL Order by ID asc"
        .Refresh
    End With
    TextSQL.Text = ""
    cmExecute.Enabled = False
    menuTool.Visible = False
    XPEngine.StartEngine
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
Sub SimpanSQLKeDatabase()
    With AdodcMain
        .Recordset.AddNew
        .Recordset.Fields(1).Value = TextSQL.Text
        .Recordset.Fields(2).Value = "(" & Day(Date) & "-" & Month(Date) & "-" & Year(Date) & ") - (" & Hour(Time) & ":" & Minute(Time) & ":" & Second(Time) & ")"
        .Recordset.Update
        .Refresh
        .Refresh
    End With
End Sub

Private Sub cmExecute_Click()
On Error GoTo HancurkanError
    With FormManage.AdodcMain
        .ConnectionString = CN.ConnectionString
        .RecordSource = TextSQL.Text
        Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain
        .Refresh
    End With
    FormManage.AturUkuranDatagrid
    SimpanSQLKeDatabase
    Unload Me
Exit Sub
HancurkanError:
    PusatError
    FormManage.AturKontrol
    TextSQL.SetFocus
End Sub

Private Sub cmHistorySQL_Click()
    If AdodcMain.Recordset.RecordCount = 0 Then
        MsgBox "Belum ada history!", vbExclamation + vbOKOnly, "Kosong"
    Else
        FormHistorySQL.Show vbModal, Me
    End If
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu menuTool
End Sub

Private Sub menuHapusText_Click()
    With TextSQL
        .Text = ""
        .SetFocus
    End With
End Sub

Private Sub textSQL_Change()
    If TextSQL.Text = "" Then
        cmExecute.Enabled = False
    Else
        cmExecute.Enabled = True
    End If
End Sub

Private Sub textSQL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu menuTool
End Sub
