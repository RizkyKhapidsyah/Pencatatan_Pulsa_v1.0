VERSION 5.00
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Begin VB.Form FormTambahReferensi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "------------------------"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormTambahReferensi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox CekTutup 
      Caption         =   "Tutup Setelah Data Disimpan"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin XPEngine.XPControl XPEngine 
      Left            =   360
      Top             =   1440
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin VB.TextBox TextKeterangan 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   4455
   End
   Begin isButton3.isButton cmSimpan 
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Icon            =   "FormTambahReferensi.frx":000C
      Style           =   0
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButton3.isButton cmBatal 
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Icon            =   "FormTambahReferensi.frx":0166
      Style           =   0
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LabelPenentu 
      AutoSize        =   -1  'True
      Caption         =   "---"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label LabelKeterangan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   180
   End
End
Attribute VB_Name = "FormTambahReferensi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    RemoveCancelMenuItem Me
    AturThema
    XPEngine.StartEngine
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

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmSimpan_Click()
    If TextKeterangan.Text = "" Then
        MsgBox "Silahkan isi referensi baru yang akan ditambahkan!", vbExclamation + vbOKOnly, ""
    Else
        X = MsgBox("Anda yakin ingin menyimpan data ini?", vbYesNo + vbQuestion, "Konfirmasi?")
        If X = vbYes Then
            Select Case LabelPenentu.Caption
                Case "1"
                    With FormBaru
                        .AdodcCmbRefJumlahPulsaDibeli.Recordset.AddNew
                        .AdodcCmbRefJumlahPulsaDibeli.Recordset.Fields(0).Value = Me.TextKeterangan.Text
                        .AdodcCmbRefJumlahPulsaDibeli.Recordset.Update
                        .AdodcCmbRefJumlahPulsaDibeli.Refresh
                        .AdodcCmbRefJumlahPulsaDibeli.Refresh
                        .IsiCMBDataLalu
                        .cmbRefJumlahPulsaDibeli.Text = Me.TextKeterangan.Text
                    End With
                    If CekTutup.Value = Checked Then Unload Me
                    Me.TextKeterangan.Text = ""
                Case "2"
                    With FormBaru
                        .AdodcCmbRefHargaServer.Recordset.AddNew
                        .AdodcCmbRefHargaServer.Recordset.Fields(0).Value = Me.TextKeterangan.Text
                        .AdodcCmbRefHargaServer.Recordset.Update
                        .AdodcCmbRefHargaServer.Refresh
                        .AdodcCmbRefHargaServer.Refresh
                        .IsiCMBDataLalu
                        .cmbRefHargaServer.Text = Me.TextKeterangan.Text
                    End With
                    If CekTutup.Value = Checked Then Unload Me
                    Me.TextKeterangan.Text = ""
                Case "3"
                    With FormBaru
                        .AdodcCmbRefHargaJual.Recordset.AddNew
                        .AdodcCmbRefHargaJual.Recordset.Fields(0).Value = Me.TextKeterangan.Text
                        .AdodcCmbRefHargaJual.Recordset.Update
                        .AdodcCmbRefHargaJual.Refresh
                        .AdodcCmbRefHargaJual.Refresh
                        .IsiCMBDataLalu
                        .CmbRefHargaJual.Text = Me.TextKeterangan.Text
                    End With
                    Me.TextKeterangan.Text = ""
                    If CekTutup.Value = Checked Then Unload Me
            End Select
        End If
    End If
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub
