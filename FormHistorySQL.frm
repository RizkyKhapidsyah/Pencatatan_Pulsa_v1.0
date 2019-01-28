VERSION 5.00
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Begin VB.Form FormHistorySQL 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "History SQL"
   ClientHeight    =   630
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3105
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormHistorySQL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   3105
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbHistorySQL 
      Height          =   375
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin XPEngine.XPControl XPEngine 
      Left            =   240
      Top             =   2520
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin isButton3.isButton cmFilter 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Style           =   6
      Caption         =   "Masukkan"
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
End
Attribute VB_Name = "FormHistorySQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub AturKontrol()
    With Me
        .cmbHistorySQL.Clear
        Do Until FormExecuteSQL.AdodcMain.Recordset.EOF
            .cmbHistorySQL.AddItem FormExecuteSQL.AdodcMain.Recordset.Fields(1).Value, 0
            FormExecuteSQL.AdodcMain.Recordset.MoveNext
        Loop
        FormExecuteSQL.AdodcMain.Refresh
        .cmbHistorySQL.ListIndex = 0
    End With
    R = SendMessageLong(cmbHistorySQL.hwnd, CB_SETDROPPEDWIDTH, 400, 0)
    XPEngine.StartEngine
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
Private Sub cmFilter_Click()
    FormExecuteSQL.TextSQL = cmbHistorySQL.Text
    Unload Me
    FormExecuteSQL.TextSQL.SetFocus
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub
