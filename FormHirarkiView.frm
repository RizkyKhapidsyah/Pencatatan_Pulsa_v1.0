VERSION 5.00
Object = "{62E1564C-1E05-44A7-A921-FD8347F324A5}#1.0#0"; "HookMenu.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "FUSIONButtons.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FormHirarkiView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hirarki View"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   4305
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormHirarkiView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   4095
      Begin VB.Label LabelInfoB12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   2160
         TabIndex        =   41
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2040
         TabIndex        =   40
         Top             =   4200
         Width           =   45
      End
      Begin VB.Label LabelInfoA12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   120
         TabIndex        =   39
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label LabelInfoB11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   2160
         TabIndex        =   38
         Top             =   3840
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2040
         TabIndex        =   37
         Top             =   3840
         Width           =   45
      End
      Begin VB.Label LabelInfoA11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         Height          =   270
         Left            =   120
         TabIndex        =   36
         Top             =   3840
         Width           =   330
      End
      Begin VB.Label LabelInfoA1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2040
         TabIndex        =   34
         Top             =   240
         Width           =   45
      End
      Begin VB.Label LabelInfoB1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   2160
         TabIndex        =   33
         Top             =   240
         Width           =   375
      End
      Begin VB.Label LabelInfoA2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2040
         TabIndex        =   31
         Top             =   600
         Width           =   45
      End
      Begin VB.Label LabelInfoB2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   2160
         TabIndex        =   30
         Top             =   600
         Width           =   375
      End
      Begin VB.Label LabelInfoA3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2040
         TabIndex        =   28
         Top             =   960
         Width           =   45
      End
      Begin VB.Label LabelInfoB3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   2160
         TabIndex        =   27
         Top             =   960
         Width           =   375
      End
      Begin VB.Label LabelInfoA4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2040
         TabIndex        =   25
         Top             =   1320
         Width           =   45
      End
      Begin VB.Label LabelInfoB4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   2160
         TabIndex        =   24
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label LabelInfoA5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2040
         TabIndex        =   22
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label LabelInfoB5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   2160
         TabIndex        =   21
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label LabelInfoA6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2040
         TabIndex        =   19
         Top             =   2040
         Width           =   45
      End
      Begin VB.Label LabelInfoB6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   2160
         TabIndex        =   18
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label LabelInfoA7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2040
         TabIndex        =   16
         Top             =   2400
         Width           =   45
      End
      Begin VB.Label LabelInfoB7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   2160
         TabIndex        =   15
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label LabelInfoA8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2040
         TabIndex        =   13
         Top             =   2760
         Width           =   45
      End
      Begin VB.Label LabelInfoB8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   2160
         TabIndex        =   12
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label LabelInfoA9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   3120
         Width           =   330
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2040
         TabIndex        =   10
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label LabelInfoB9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   2160
         TabIndex        =   9
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label LabelInfoA10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2040
         TabIndex        =   7
         Top             =   3480
         Width           =   45
      End
      Begin VB.Label LabelInfoB10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   270
         Left            =   2160
         TabIndex        =   6
         Top             =   3480
         Width           =   375
      End
   End
   Begin VB.TextBox textData 
      Height          =   480
      Left            =   1125
      TabIndex        =   4
      Top             =   4695
      Width           =   2055
   End
   Begin XPEngine.XPControl XPEngine 
      Left            =   720
      Top             =   6120
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin KewlButtonz.KewlButtons cmDataAwal 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4680
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
      MICON           =   "FormHirarkiView.frx":014A
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
      Left            =   615
      TabIndex        =   1
      Top             =   4680
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
      MICON           =   "FormHirarkiView.frx":0166
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
      Left            =   3195
      TabIndex        =   2
      Top             =   4680
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
      MICON           =   "FormHirarkiView.frx":0182
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
      Left            =   3690
      TabIndex        =   3
      Top             =   4680
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
      MICON           =   "FormHirarkiView.frx":019E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   0
      Top             =   0
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   1
      Bmp:1           =   "FormHirarkiView.frx":01BA
      Key:1           =   "#menuSimpan"
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
   Begin VB.Menu menuTool 
      Caption         =   "Tool"
      Begin VB.Menu menuSimpan 
         Caption         =   "Simpan"
      End
   End
End
Attribute VB_Name = "FormHirarkiView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub MasukkanDatabaseKeDataLabel()
    With Me
        .LabelInfoA1.Caption = "Tanggal"
        .LabelInfoA2.Caption = FormUtama.AdodcMain.Recordset.Fields(4).Name
        .LabelInfoA3.Caption = FormUtama.AdodcMain.Recordset.Fields(5).Name
        .LabelInfoA4.Caption = FormUtama.AdodcMain.Recordset.Fields(6).Name
        .LabelInfoA5.Caption = FormUtama.AdodcMain.Recordset.Fields(7).Name
        .LabelInfoA6.Caption = FormUtama.AdodcMain.Recordset.Fields(8).Name
        .LabelInfoA7.Caption = FormUtama.AdodcMain.Recordset.Fields(9).Name
        .LabelInfoA8.Caption = FormUtama.AdodcMain.Recordset.Fields(10).Name
        .LabelInfoA9.Caption = FormUtama.AdodcMain.Recordset.Fields(11).Name
        .LabelInfoA10.Caption = FormUtama.AdodcMain.Recordset.Fields(12).Name
        .LabelInfoA11.Caption = FormUtama.AdodcMain.Recordset.Fields(13).Name
        .LabelInfoA12.Caption = FormUtama.AdodcMain.Recordset.Fields(14).Name
        .LabelInfoB1.Caption = FormUtama.AdodcMain.Recordset.Fields(0).Value & ", " & FormUtama.AdodcMain.Recordset.Fields(1).Value & " - " & FormUtama.AdodcMain.Recordset.Fields(2).Value & " - " & FormUtama.AdodcMain.Recordset.Fields(3).Value
        .LabelInfoB2.Caption = FormUtama.AdodcMain.Recordset.Fields(4).Value
        .LabelInfoB3.Caption = FormUtama.AdodcMain.Recordset.Fields(5).Value
        .LabelInfoB4.Caption = FormUtama.AdodcMain.Recordset.Fields(6).Value
        .LabelInfoB5.Caption = FormUtama.AdodcMain.Recordset.Fields(7).Value
        .LabelInfoB6.Caption = FormUtama.AdodcMain.Recordset.Fields(8).Value
        .LabelInfoB7.Caption = FormUtama.AdodcMain.Recordset.Fields(9).Value
        .LabelInfoB8.Caption = FormUtama.AdodcMain.Recordset.Fields(10).Value
        .LabelInfoB9.Caption = FormUtama.AdodcMain.Recordset.Fields(11).Value
        .LabelInfoB10.Caption = FormUtama.AdodcMain.Recordset.Fields(12).Value
        .LabelInfoB11.Caption = FormUtama.AdodcMain.Recordset.Fields(13).Value
        .LabelInfoB12.Caption = FormUtama.AdodcMain.Recordset.Fields(14).Value
    End With
    With textData
        .BackColor = Me.BackColor
        .Locked = True
        .Text = "Data ke '" & FormUtama.AdodcMain.Recordset.AbsolutePosition & "' dari '" & FormUtama.AdodcMain.Recordset.RecordCount & "' data"
        .Alignment = 2
    End With
    menuTool.Visible = False
    XPEngine.StartEngine
End Sub

Private Sub cmDataAkhir_Click()
    FormUtama.AdodcMain.Recordset.MoveLast
    MasukkanDatabaseKeDataLabel
End Sub

Private Sub cmDataAwal_Click()
    FormUtama.AdodcMain.Recordset.MoveFirst
    MasukkanDatabaseKeDataLabel
End Sub

Private Sub cmDataSebelumnya_Click()
    FormUtama.AdodcMain.Recordset.MovePrevious
    If FormUtama.AdodcMain.Recordset.BOF Then FormUtama.AdodcMain.Recordset.MoveLast
    MasukkanDatabaseKeDataLabel
End Sub

Private Sub cmDataSelanjutnya_Click()
    FormUtama.AdodcMain.Recordset.MoveNext
    If FormUtama.AdodcMain.Recordset.EOF Then FormUtama.AdodcMain.Recordset.MoveFirst
    MasukkanDatabaseKeDataLabel
End Sub

Private Sub Form_Load()
    MasukkanDatabaseKeDataLabel
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu menuTool
End Sub

Private Sub menuSimpan_Click()
    On Error GoTo ErrorHandler
    CommonDialog1.Filter = "All Files (*.*)|*.*|RikySoft Catatan Files (*.rcf)|*.rcf|Text Files (*.txt)|*.txt"
    DefaultFormat
    CommonDialog1.ShowSave
    CommonDialog1.FileName = CommonDialog1.FileName
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
    Print #iFile, LabelInfoA1.Caption & " : " & LabelInfoB1.Caption
    Print #iFile, LabelInfoA2.Caption & " : " & LabelInfoB2.Caption
    Print #iFile, LabelInfoA3.Caption & " : " & LabelInfoB3.Caption
    Print #iFile, LabelInfoA4.Caption & " : " & LabelInfoB4.Caption
    Print #iFile, LabelInfoA5.Caption & " : " & LabelInfoB5.Caption
    Print #iFile, LabelInfoA6.Caption & " : " & LabelInfoB6.Caption
    Print #iFile, LabelInfoA7.Caption & " : " & LabelInfoB7.Caption
    Print #iFile, LabelInfoA8.Caption & " : " & LabelInfoB8.Caption
    Print #iFile, LabelInfoA9.Caption & " : " & LabelInfoB9.Caption
    Print #iFile, LabelInfoA10.Caption & " : " & LabelInfoB10.Caption
    Print #iFile, LabelInfoA11.Caption & " : " & LabelInfoB11.Caption
    Print #iFile, LabelInfoA12.Caption & " : " & LabelInfoB12.Caption
    SaveFileFromTB = True
ErrorHandler:
    Close #iFile
End Sub
