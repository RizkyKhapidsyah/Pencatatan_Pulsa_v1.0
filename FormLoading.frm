VERSION 5.00
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormLoading 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   525
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4995
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FormLoading.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   525
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Left            =   3960
      Top             =   240
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   2640
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   120
      Top             =   2400
   End
   Begin MSComctlLib.ProgressBar Proses 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin isButton3.isButton CmBatalkan 
      Height          =   390
      Left            =   4200
      TabIndex        =   2
      Top             =   60
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   688
      Icon            =   "FormLoading.frx":000C
      Style           =   0
      Caption         =   "&Batalkan"
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
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   180
   End
End
Attribute VB_Name = "FormLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CmBatalkan_Click()
    CmBatalkan.Enabled = False
    With Timer1
        .Interval = 0
        .Enabled = False
    End With
    With Timer2
        .Enabled = True
        .Interval = 4000
    End With
    With Timer3
        .Enabled = True
        .Interval = 3500
    End With
End Sub

Private Sub Form_Load()
    Proses.Value = 1
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
Private Sub Timer1_Timer()
    On Error GoTo Ero
    Proses.Value = Proses.Value + 1
    LabelPersen.Caption = Proses.Value & "%"
    If Proses.Value = 100 Then
    CmBatalkan.Enabled = False
   
    Dim xlApp As New Excel.Application
    
    
    With xlApp
    
    .Workbooks.Add
    
    'judul
    .Range("A1").Value = "RikySoft - Pencatatan Pulsa"
    .Range("A1").Select
    .Selection.Font.Bold = True
    .Selection.Font.Size = 16
    
    'kolom
    .Range("A2").Value = FormManage.AdodcMain.Recordset.Fields(0).Name
    .Range("B2").Value = FormManage.AdodcMain.Recordset.Fields(1).Name
    .Range("C2").Value = FormManage.AdodcMain.Recordset.Fields(2).Name
    .Range("D2").Value = FormManage.AdodcMain.Recordset.Fields(3).Name
    .Range("E2").Value = FormManage.AdodcMain.Recordset.Fields(4).Name
    .Range("F2").Value = FormManage.AdodcMain.Recordset.Fields(5).Name
    .Range("G2").Value = FormManage.AdodcMain.Recordset.Fields(6).Name
    .Range("H2").Value = FormManage.AdodcMain.Recordset.Fields(7).Name
    .Range("I2").Value = FormManage.AdodcMain.Recordset.Fields(8).Name
    .Range("J2").Value = FormManage.AdodcMain.Recordset.Fields(9).Name
    .Range("K2").Value = FormManage.AdodcMain.Recordset.Fields(10).Name
    .Range("L2").Value = FormManage.AdodcMain.Recordset.Fields(11).Name
    .Range("M2").Value = FormManage.AdodcMain.Recordset.Fields(12).Name
    .Range("N2").Value = FormManage.AdodcMain.Recordset.Fields(13).Name
    .Range("O2").Value = FormManage.AdodcMain.Recordset.Fields(14).Name
    .Range("A2:O2").Select
    .Selection.Font.Bold = True
    .Selection.HorizontalAlignment = xlCenter
    
    'data
    FormManage.AdodcMain.Recordset.MoveFirst
    For X = 1 To FormManage.AdodcMain.Recordset.RecordCount
        .Range("A" & CStr(X + 2)).Value = FormManage.AdodcMain.Recordset.Fields(0).Value
        .Range("B" & CStr(X + 2)).Value = FormManage.AdodcMain.Recordset.Fields(1).Value
        .Range("C" & CStr(X + 2)).Value = FormManage.AdodcMain.Recordset.Fields(2).Value
        .Range("D" & CStr(X + 2)).Value = FormManage.AdodcMain.Recordset.Fields(3).Value
        .Range("E" & CStr(X + 2)).Value = FormManage.AdodcMain.Recordset.Fields(4).Value
        .Range("F" & CStr(X + 2)).Value = FormManage.AdodcMain.Recordset.Fields(5).Value
        .Range("G" & CStr(X + 2)).Value = FormManage.AdodcMain.Recordset.Fields(6).Value
        .Range("H" & CStr(X + 2)).Value = FormManage.AdodcMain.Recordset.Fields(7).Value
        .Range("I" & CStr(X + 2)).Value = FormManage.AdodcMain.Recordset.Fields(8).Value
        .Range("J" & CStr(X + 2)).Value = FormManage.AdodcMain.Recordset.Fields(9).Value
        .Range("K" & CStr(X + 2)).Value = FormManage.AdodcMain.Recordset.Fields(10).Value
        .Range("L" & CStr(X + 2)).Value = FormManage.AdodcMain.Recordset.Fields(11).Value
        .Range("M" & CStr(X + 2)).Value = FormManage.AdodcMain.Recordset.Fields(12).Value
        .Range("N" & CStr(X + 2)).Value = FormManage.AdodcMain.Recordset.Fields(13).Value
        .Range("O" & CStr(X + 2)).Value = FormManage.AdodcMain.Recordset.Fields(14).Value
        FormManage.AdodcMain.Recordset.MoveNext
    Next
    
    'membuat list
    xlApp.ActiveSheet.ListObjects.Add , xlApp.Range("A2:O" & CStr(FormManage.AdodcMain.Recordset.RecordCount + 2)), , xlYes
    
    .Range("A1").Select
    .Visible = True
    
    End With
    
    
        Timer1.Enabled = False
        Unload Me
    End If
    Exit Sub
    
Ero:
    MsgBox Err.Description
    xlApp.ActiveWorkbook.Close False
End Sub

Private Sub Timer2_Timer()
    Unload Me
End Sub

Private Sub Timer3_Timer()
    Proses.Value = 100
    LabelPersen.Caption = Proses.Value & "%"
End Sub
