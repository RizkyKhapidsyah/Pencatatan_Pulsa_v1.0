VERSION 5.00
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Begin VB.Form FormMultiFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi Filter"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormMultiFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox LBMultiFilter 
      Height          =   2490
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin XPEngine.XPControl XPEngine 
      Left            =   240
      Top             =   2640
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filter Hanya Kolom :"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1290
   End
End
Attribute VB_Name = "FormMultiFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    XPEngine.StartEngine
    With LBMultiFilter
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
        .Selected (9)
        
        .ListIndex = 0
    End With
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub
