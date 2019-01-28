VERSION 5.00
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Begin VB.Form FormTambahEditListTelepon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "-------------"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3720
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormTambahEditListTelepon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPEngine.XPControl XPEngine 
      Left            =   240
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
      Width           =   3495
      Begin VB.TextBox textNomorTelepon 
         Height          =   390
         Left            =   1560
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox textNama 
         Height          =   390
         Left            =   1560
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1440
         TabIndex        =   5
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor_Telepon"
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   360
      End
   End
   Begin isButton3.isButton cmSimpan 
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormTambahEditListTelepon.frx":57E2
      Style           =   6
      Caption         =   "&Simpan"
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
      Left            =   2400
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormTambahEditListTelepon.frx":593C
      Style           =   6
      Caption         =   "&Batal"
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
Attribute VB_Name = "FormTambahEditListTelepon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub AturKontrol()
    RemoveCancelMenuItem Me
    XPEngine.StartEngine
    Reset
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
Sub Reset()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .Text = ""
                .MaxLength = 254
            End With
        End If
    Next
End Sub

Private Sub cmSimpan_Click()
On Error GoTo HancurkanError
    If textNama.Text = "" Then
        MsgBox "Silahkan isi Nama !", vbExclamation + vbOKOnly, ""
        textNama.SetFocus
    ElseIf textNomorTelepon.Text = "" Then
        MsgBox "Silahkan isi Nomor Telepon!", vbExclamation + vbOKOnly, ""
        textNomorTelepon.SetFocus
    Else
        If cmSimpan.Caption = "&Simpan" Then
                With FormListTelepon
                    .AdodcMain.Recordset.AddNew
                    .AdodcMain.Recordset.Fields(0).Value = textNama.Text
                    .AdodcMain.Recordset.Fields(1).Value = textNomorTelepon.Text
                    .AdodcMain.Recordset.Update
                    .AdodcMain.Refresh
                    .AturKontrol
                    .AdodcMain.Recordset.MoveLast
                End With
                Reset
                cmTutup.Caption = "&Tutup"
                textNama.SetFocus
        ElseIf cmSimpan.Caption = "&Update" Then
            With FormListTelepon
                .AdodcMain.Recordset.Delete
                .AdodcMain.Recordset.AddNew
                .AdodcMain.Recordset.Fields(0).Value = textNama.Text
                .AdodcMain.Recordset.Fields(1).Value = textNomorTelepon.Text
                .AdodcMain.Recordset.Update
                .AdodcMain.Refresh
                .AturKontrol
                .AdodcMain.Recordset.MoveLast
            End With
            Unload Me
        End If
    End If
Exit Sub
HancurkanError:
    PusatError
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub
