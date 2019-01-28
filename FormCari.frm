VERSION 5.00
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Begin VB.Form FormCari 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cari Data"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5040
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormCari.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4815
      Begin VB.ComboBox cmbCariBerdasarkan 
         Height          =   390
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox textKriteria 
         Height          =   390
         Left            =   2040
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cari Berdasarkan"
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kriteria"
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1920
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
   Begin isButton3.isButton cmCari 
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormCari.frx":0582
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
      Left            =   3720
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormCari.frx":0B14
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
Attribute VB_Name = "FormCari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    With cmbCariBerdasarkan
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
    On Error GoTo HancurkanError
    If textKriteria.Text = "" Then
        MsgBox "Silahkan isi kriteria yang akan dicari!", vbExclamation + vbOKOnly, ""
        textKriteria.SetFocus
    Else
        FormManage.AdodcMain.Refresh
            With FormManage.AdodcMain.Recordset
                Select Case Me.cmbCariBerdasarkan.ListIndex
                Case Is = 0
                    .Find "Hari = '" & textKriteria.Text & "'"
                Case Is = 1
                    .Find "Tanggal = '" & textKriteria.Text & "'"
                Case Is = 2
                    .Find "Bulan = '" & textKriteria.Text & "'"
                Case Is = 3
                    .Find "Tahun = '" & textKriteria.Text & "'"
                Case Is = 4
                    .Find "Penerima_Pulsa = '" & textKriteria.Text & "'"
                Case Is = 5
                    .Find "Nama_Penerima = '" & textKriteria.Text & "'"
                Case Is = 6
                    .Find "Provider = '" & textKriteria.Text & "'"
                Case Is = 7
                    .Find "Jenis_Provider = '" & textKriteria.Text & "'"
                Case Is = 8
                    .Find "Jumlah_Pulsa_Dibeli = '" & textKriteria.Text & "'"
                Case Is = 9
                    .Find "Harga_Server = '" & textKriteria.Text & "'"
                Case Is = 10
                    .Find "Harga_Jual = '" & textKriteria.Text & "'"
                Case Is = 11
                    .Find "Uang_Bayar = '" & textKriteria.Text & "'"
                Case Is = 12
                    .Find "Tanggal_Bayar = '" & textKriteria.Text & "'"
                Case Is = 13
                    .Find "Laba = '" & textKriteria.Text & "'"
                Case Is = 14
                    .Find "Status_Bayar = '" & textKriteria.Text & "'"
                End Select
                '=============================================================================
                If .EOF Then
                    MsgBox ("Data tidak ditemukan"), vbExclamation + vbOKOnly, ""
                    With FormManage
                        .AdodcMain.Refresh
                        .cmRefresh_Click
                    End With
                Else
                    Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                    FormManage.AturUkuranDatagrid
                            MsgBox FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(8).Name & " : " & FormManage.AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(9).Name & " : " & FormManage.AdodcMain.Recordset.Fields(9).Value & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(10).Name & " : " & FormManage.AdodcMain.Recordset.Fields(10).Value & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(11).Name & " : " & FormManage.AdodcMain.Recordset.Fields(11).Value & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(12).Name & " : " & FormManage.AdodcMain.Recordset.Fields(12).Value & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(13).Name & " : " & FormManage.AdodcMain.Recordset.Fields(13).Value & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(14).Name & " : " & FormManage.AdodcMain.Recordset.Fields(14).Value, vbInformation + vbOKOnly, "Data Ditemukan!"
                            With FormHasilPencarian
                                .TextDitemukan.Text = FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                                        FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                                        FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                                        FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                                        FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                                        FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                                        FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                                        FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                                                        FormManage.AdodcMain.Recordset.Fields(8).Name & " : " & FormManage.AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                                                        FormManage.AdodcMain.Recordset.Fields(9).Name & " : " & FormManage.AdodcMain.Recordset.Fields(9).Value & vbCrLf & _
                                                        FormManage.AdodcMain.Recordset.Fields(10).Name & " : " & FormManage.AdodcMain.Recordset.Fields(10).Value & vbCrLf & _
                                                        FormManage.AdodcMain.Recordset.Fields(11).Name & " : " & FormManage.AdodcMain.Recordset.Fields(11).Value & vbCrLf & _
                                                        FormManage.AdodcMain.Recordset.Fields(12).Name & " : " & FormManage.AdodcMain.Recordset.Fields(12).Value & vbCrLf & _
                                                        FormManage.AdodcMain.Recordset.Fields(13).Name & " : " & FormManage.AdodcMain.Recordset.Fields(13).Value & vbCrLf & _
                                                        FormManage.AdodcMain.Recordset.Fields(14).Name & " : " & FormManage.AdodcMain.Recordset.Fields(14).Value
                                .Show vbModal, Me
                            End With
                                Set FormManage.DataGrid1.DataSource = FormManage.AdodcMain.Recordset
                End If
            End With
    End If
Exit Sub
HancurkanError:
    PusatError
    With FormManage
        .AdodcMain.Refresh
        .cmRefresh_Click
    End With
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    FormManage.AturUkuranDatagrid
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub
