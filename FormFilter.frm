VERSION 5.00
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Begin VB.Form FormFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter Data"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.ComboBox cmbMode 
         Height          =   390
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox cmbFilterBerdasarkan 
         Height          =   390
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
      Begin isButton3.isButton cmFilter 
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Icon            =   "FormFilter.frx":014A
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   7
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mode"
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter Berdasarkan"
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1200
      End
   End
   Begin XPEngine.XPControl XPEngine 
      Left            =   840
      Top             =   3000
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
End
Attribute VB_Name = "FormFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    XPEngine.StartEngine
    With cmbFilterBerdasarkan
        .Clear
        .AddItem FormManage.AdodcMain.Recordset.Fields(0).Name, 0
        .AddItem FormManage.AdodcMain.Recordset.Fields(1).Name, 1
        .AddItem FormManage.AdodcMain.Recordset.Fields(2).Name, 2
        .AddItem FormManage.AdodcMain.Recordset.Fields(3).Name, 3
        .AddItem FormManage.AdodcMain.Recordset.Fields(4).Name, 4
        .AddItem FormManage.AdodcMain.Recordset.Fields(5).Name, 5
        .AddItem FormManage.AdodcMain.Recordset.Fields(6).Name, 6
        .AddItem FormManage.AdodcMain.Recordset.Fields(7).Name, 7
        .AddItem FormManage.AdodcMain.Recordset.Fields(8).Name, 8
        .AddItem FormManage.AdodcMain.Recordset.Fields(9).Name, 9
        .AddItem FormManage.AdodcMain.Recordset.Fields(10).Name, 10
        .AddItem FormManage.AdodcMain.Recordset.Fields(11).Name, 11
        .AddItem FormManage.AdodcMain.Recordset.Fields(12).Name, 12
        .AddItem FormManage.AdodcMain.Recordset.Fields(13).Name, 13
        .AddItem FormManage.AdodcMain.Recordset.Fields(14).Name, 14
        .ListIndex = 0
    End With
    With cmbMode
        .Clear
        .AddItem "Asc", 0
        .AddItem "Desc", 1
        .ListIndex = 0
    End With
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
    With FormManage
        .AdodcMain.RecordSource = "Select " & cmbFilterBerdasarkan.Text & " from TablePulsa order by " & cmbFilterBerdasarkan.Text & " " & cmbMode.Text & ";"
        .AdodcMain.Refresh
        .cmEdit.Enabled = False
        .cmCari.Enabled = False
        .cmSorot.Enabled = False
        .cmHapus.Enabled = False
        .cmEskpor.Enabled = False
    End With
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

