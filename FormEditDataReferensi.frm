VERSION 5.00
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Begin VB.Form FormEditDataReferensi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "--------------"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3945
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormEditDataReferensi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox textEdit 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   3255
   End
   Begin XPEngine.XPControl XPEngine 
      Left            =   840
      Top             =   2280
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin isButton3.isButton cmUpdate 
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormEditDataReferensi.frx":014A
      Style           =   0
      Caption         =   "&Update"
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
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormEditDataReferensi.frx":02A4
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
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButton3.isButton cmKosongkanText 
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Icon            =   "FormEditDataReferensi.frx":03FE
      Style           =   0
      Caption         =   "'"
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
   Begin VB.Label LabelEdit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   315
   End
End
Attribute VB_Name = "FormEditDataReferensi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmKosongkanText_Click()
    With textEdit
        .Text = ""
        .SetFocus
    End With
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub cmUpdate_Click()
    If textEdit.Text = "" Then
        MsgBox "Silahkan input nilai yang diinginkan", vbExclamation + vbOKOnly, ""
        textEdit.SetFocus
    Else
        X = MsgBox("Anda yakin ingin memperbarui data dengan nilai ini?", vbQuestion + vbYesNo, "Konfirmasi?")
        If X = vbYes Then
            With FormDataReferensi
                .AdodcMain.Recordset.Delete
                .AdodcMain.Recordset.AddNew
                .AdodcMain.Recordset.Fields(0).Value = textEdit.Text
                .AdodcMain.Recordset.Update
                .AdodcMain.Refresh
                .AdodcMain.Refresh
                .DataGrid1.Columns.Item(0).Width = 3000
            End With
        Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    AturThema
    RemoveCancelMenuItem Me
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
