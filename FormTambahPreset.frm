VERSION 5.00
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Begin VB.Form FormTambahPreset 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tambah Preset"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5280
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormTambahPreset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox CekTutupForm 
      Caption         =   "Tutup Setelah data baru disimpan"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   5055
      Begin VB.TextBox textNamaPreset 
         Height          =   390
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "FormTambahPreset.frx":000C
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox textSyntaxSQL 
         Height          =   975
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "FormTambahPreset.frx":0012
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Preset"
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Syntax SQL"
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1560
         TabIndex        =   5
         Top             =   720
         Width           =   45
      End
   End
   Begin XPEngine.XPControl XPEngine 
      Left            =   240
      Top             =   2640
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin isButton3.isButton cmBatal 
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormTambahPreset.frx":0018
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
   Begin isButton3.isButton cmSimpan 
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormTambahPreset.frx":0172
      Style           =   6
      Caption         =   "   &Simpan"
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
Attribute VB_Name = "FormTambahPreset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    Reset
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
Sub Reset()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .Text = ""
                .MaxLength = 254
            End With
        End If
    Next
    textNamaPreset.ToolTipText = "Nama yang akan ditampilan untuk preset"
    textSyntaxSQL.ToolTipText = "Perhatikan syntax SQL dengan benar. " & vbCrLf & _
                                "Syntax SQL yang salah tidak akan dapat diperoses oleh sistem. " & vbCrLf & _
                                "Dan perhatikan juga logika syntax. Syntax yang salah logika " & vbCrLf & _
                                "dapat mempengaruhi urutan dan pemfilteran data yang tidak diinginkan"
End Sub

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmSimpan_Click()
    If textNamaPreset.Text = "" Then
        MsgBox "Silahkan isi nama preset yang akan ditampilan!", vbExclamation + vbOKOnly, ""
        textNamaPreset.SetFocus
    ElseIf textSyntaxSQL.Text = "" Then
        MsgBox "Silahkan isi syntax SQL yang dipakai untuk mengolah data!", vbExclamation + vbOKOnly, ""
        textSyntaxSQL.SetFocus
    Else
        If cmSimpan.Caption = "   &Simpan" Then
            Pesan = MsgBox("Anda yakin ingin menyimpan preset ini?", vbQuestion + vbYesNo, "Konfirmasi")
            If Pesan = vbYes Then
                With FormPreset
                    .AdodcMain.Recordset.AddNew
                    .AdodcMain.Recordset.Fields(0).Value = textNamaPreset.Text
                    .AdodcMain.Recordset.Fields(1).Value = textSyntaxSQL.Text
                    .AdodcMain.Recordset.Update
                    .AdodcMain.Refresh
                    .AdodcMain.Refresh
                    .AturKontrol
                    .cmbPreset.Text = Me.textNamaPreset.Text
                End With
                cmBatal.Caption = "&Tutup"
                AturKontrol
                FormPreset.AturKontrol
                If CekTutupForm.Value = Checked Then Unload Me
            End If
        ElseIf cmSimpan.Caption = "  &Update" Then
            Pesan = MsgBox("Anda yakin ingin memperbarui preset ini?", vbQuestion + vbYesNo, "Konfirmasi")
            If Pesan = vbYes Then
                With FormAturPreset
                    .AdodcMain.Recordset.Delete
                    .AdodcMain.Recordset.AddNew
                    .AdodcMain.Recordset.Fields(0).Value = textNamaPreset.Text
                    .AdodcMain.Recordset.Fields(1).Value = textSyntaxSQL.Text
                    .AdodcMain.Recordset.Update
                    .AdodcMain.Refresh
                    .AdodcMain.Refresh
                    .AturKontrol
                End With
                FormPreset.AturKontrol
                If CekTutupForm.Value = Checked Then Unload Me
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub
