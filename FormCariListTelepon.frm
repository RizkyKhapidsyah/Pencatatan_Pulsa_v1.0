VERSION 5.00
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Begin VB.Form FormCariListTelepon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cari"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4185
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormCariListTelepon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPEngine.XPControl XPEngine 
      Left            =   120
      Top             =   1440
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.TextBox textKriteria 
         Height          =   390
         Left            =   2040
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox cmbCariBerdasarkan 
         Height          =   375
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   6
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kriteria"
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cari Berdasarkan"
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1155
      End
   End
   Begin isButton3.isButton cmCari 
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormCariListTelepon.frx":0582
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
   Begin isButton3.isButton cmTutup 
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormCariListTelepon.frx":0B14
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
End
Attribute VB_Name = "FormCariListTelepon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    With cmbCariBerdasarkan
        .Clear
        .AddItem FormListTelepon.AdodcMain.Recordset.Fields(0).Name, 0
        .AddItem FormListTelepon.AdodcMain.Recordset.Fields(1).Name, 1
        .ListIndex = 0
    End With
    RemoveCancelMenuItem Me
    textKriteria.Text = ""
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
Private Sub cmCari_Click()
    If textKriteria.Text = "" Then
        MsgBox "Silahkan isi kriteria yang akan dicari!", vbExclamation + vbOKOnly, ""
        textKriteria.SetFocus
    Else
        FormListTelepon.AdodcMain.Refresh
            With FormListTelepon.AdodcMain.Recordset
                Select Case Me.cmbCariBerdasarkan.ListIndex
                Case Is = 0
                    .Find "Nama = '" & textKriteria.Text & "'"
                Case Is = 1
                    .Find "Nomor_Telepon = '" & textKriteria.Text & "'"
                End Select
                '=============================================================================
                If .EOF Then
                    MsgBox ("Data tidak ditemukan"), vbExclamation + vbOKOnly, ""
                    FormListTelepon.AturUkuranDatagrid
                Else
                    Set FormListTelepon.DataGrid1.DataSource = FormListTelepon.AdodcMain.Recordset
                    FormListTelepon.AturUkuranDatagrid
                            MsgBox FormListTelepon.AdodcMain.Recordset.Fields(0).Name & " : " & FormListTelepon.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                    FormListTelepon.AdodcMain.Recordset.Fields(1).Name & " : " & FormListTelepon.AdodcMain.Recordset.Fields(1).Value, vbInformation + vbOKOnly, "Data Ditemukan!"
                            With FormHasilPencarian
                                .TextDitemukan.Text = FormListTelepon.AdodcMain.Recordset.Fields(0).Name & " : " & FormListTelepon.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                                        FormListTelepon.AdodcMain.Recordset.Fields(1).Name & " : " & FormListTelepon.AdodcMain.Recordset.Fields(1).Value
                                .Show vbModal, Me
                            End With
                                Set FormListTelepon.DataGrid1.DataSource = FormListTelepon.AdodcMain.Recordset
                End If
            End With
    End If
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    FormListTelepon.AturUkuranDatagrid
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub


