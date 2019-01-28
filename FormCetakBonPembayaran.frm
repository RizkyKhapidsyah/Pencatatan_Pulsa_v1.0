VERSION 5.00
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form FormCetakBonPembayaran 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cetak Bon"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5385
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormCetakBonPembayaran.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar Proses 
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   840
      Top             =   3120
   End
   Begin VB.TextBox textPreviewBon 
      Height          =   6015
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Text            =   "FormCetakBonPembayaran.frx":014A
      Top             =   120
      Width           =   3975
   End
   Begin XPEngine.XPControl XPEngine 
      Left            =   720
      Top             =   3120
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin isButton3.isButton cmBatal 
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   6240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormCetakBonPembayaran.frx":0150
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
   Begin isButton3.isButton cmCetak 
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   6240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormCetakBonPembayaran.frx":02AA
      Style           =   6
      Caption         =   "&Cetak"
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
   Begin VB.Label LabelPersen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   180
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview Bon"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   870
   End
End
Attribute VB_Name = "FormCetakBonPembayaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    XPEngine.StartEngine
    RemoveCancelMenuItem Me
    With Me
        .textPreviewBon.Locked = True
    End With
    AturThema
    For Each Objek In Me
        If TypeName(Objek) = "Label" Or TypeName(Objek) = "isButton" Or TypeName(Objek) = "TextBox" Or TypeName(Objek) = "ComboBox" Then
            Objek.Enabled = False
        End If
    Next
    Proses.Value = 1
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


Private Sub cmCetak_Click()
    On Error GoTo ErrHandler
    Dim BeginPage, EndPage, NumCopies, i
    CommonDialog1.CancelError = True
    CommonDialog1.ShowPrinter
    BeginPage = CommonDialog1.FromPage
    EndPage = CommonDialog1.ToPage
    NumCopies = CommonDialog1.Copies
    For i = 1 To NumCopies
    Printer.Print textPreviewBon.Text
    Next i
    Exit Sub
ErrHandler:
    Exit Sub
End Sub

Private Sub Timer1_Timer()
    Proses.Value = Proses.Value + 1
    LabelPersen.Caption = Proses.Value & "%"
    With FormCetakBonPembayaran
        If Proses.Value = 2 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value
        ElseIf Proses.Value = 4 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value
        ElseIf Proses.Value = 10 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " ("
        ElseIf Proses.Value = 13 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value
        ElseIf Proses.Value = 21 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - "
        ElseIf Proses.Value = 29 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value
        ElseIf Proses.Value = 30 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-"
        ElseIf Proses.Value = 35 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value
        ElseIf Proses.Value = 38 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "=================================="
        ElseIf Proses.Value = 40 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name
        ElseIf Proses.Value = 41 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : "
        ElseIf Proses.Value = 43 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value
        ElseIf Proses.Value = 49 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name
        ElseIf Proses.Value = 50 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : "
        ElseIf Proses.Value = 51 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value
        ElseIf Proses.Value = 52 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(2).Name
        ElseIf Proses.Value = 53 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(2).Name & " : "
        ElseIf Proses.Value = 55 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value
        ElseIf Proses.Value = 59 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(3).Name
        ElseIf Proses.Value = 60 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(3).Name & " : "
        ElseIf Proses.Value = 64 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value
        ElseIf Proses.Value = 66 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(4).Name
        ElseIf Proses.Value = 69 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(4).Name & " : "
        ElseIf Proses.Value = 71 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value
        ElseIf Proses.Value = 72 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(5).Name
        ElseIf Proses.Value = 73 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(5).Name & " : "
        ElseIf Proses.Value = 77 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value
        ElseIf Proses.Value = 80 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(6).Name
        ElseIf Proses.Value = 82 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value
        ElseIf Proses.Value = 85 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(8).Name & " : " & FormManage.AdodcMain.Recordset.Fields(8).Value
        ElseIf Proses.Value = 86 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(8).Name & " : " & FormManage.AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(9).Name
        ElseIf Proses.Value = 87 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(8).Name & " : " & FormManage.AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(9).Name & " : " & FormManage.AdodcMain.Recordset.Fields(9).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(10).Name
        ElseIf Proses.Value = 90 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(8).Name & " : " & FormManage.AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(9).Name & " : " & FormManage.AdodcMain.Recordset.Fields(9).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(10).Name & " : "
        ElseIf Proses.Value = 92 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(1).Name & " : " & FormManage.AdodcMain.Recordset.Fields(1).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(2).Name & " : " & FormManage.AdodcMain.Recordset.Fields(2).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(3).Name & " : " & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(4).Name & " : " & FormManage.AdodcMain.Recordset.Fields(4).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(5).Name & " : " & FormManage.AdodcMain.Recordset.Fields(5).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(6).Name & " : " & FormManage.AdodcMain.Recordset.Fields(6).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(7).Name & " : " & FormManage.AdodcMain.Recordset.Fields(7).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(8).Name & " : " & FormManage.AdodcMain.Recordset.Fields(8).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(9).Name & " : " & FormManage.AdodcMain.Recordset.Fields(9).Value & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(10).Name & " : " & FormManage.AdodcMain.Recordset.Fields(10).Value
        ElseIf Proses.Value = 93 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
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
                                   FormManage.AdodcMain.Recordset.Fields(13).Name & " : "
        ElseIf Proses.Value = 96 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
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
                                   FormManage.AdodcMain.Recordset.Fields(13).Name & " : " & FormManage.AdodcMain.Recordset.Fields(13).Value
        ElseIf Proses.Value = 97 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
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
                                   FormManage.AdodcMain.Recordset.Fields(14).Name
        ElseIf Proses.Value = 99 Then
            .textPreviewBon.Text = "==================================" & vbCrLf & _
                                    FormManage.AdodcMain.Recordset.Fields(5).Value & " (" & FormManage.AdodcMain.Recordset.Fields(4).Value & ") - " & FormManage.AdodcMain.Recordset.Fields(0).Value & ", " & FormManage.AdodcMain.Recordset.Fields(1).Value & "-" & FormManage.AdodcMain.Recordset.Fields(2).Value & "-" & FormManage.AdodcMain.Recordset.Fields(3).Value & vbCrLf & _
                                   "==================================" & vbCrLf & vbCrLf & _
                                   FormManage.AdodcMain.Recordset.Fields(0).Name & " : " & FormManage.AdodcMain.Recordset.Fields(0).Value & vbCrLf & _
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
                                   FormManage.AdodcMain.Recordset.Fields(14).Name & " : " & FormManage.AdodcMain.Recordset.Fields(14).Value & vbCrLf & _
                                   "=================================="
        ElseIf Proses.Value = 100 Then
            Timer1.Enabled = False
            With Proses
                .Enabled = False
                .Visible = False
            End With
            For Each Objek In Me
                If TypeName(Objek) = "Label" Or TypeName(Objek) = "isButton" Or TypeName(Objek) = "TextBox" Or TypeName(Objek) = "ComboBox" Then
                    Objek.Enabled = True
                End If
            Next
            LabelPersen.Visible = False
        End If
    End With
End Sub

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub
