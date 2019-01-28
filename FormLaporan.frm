VERSION 5.00
Object = "{02353968-C1C9-4E0A-88D3-18759BDC60FE}#1.0#0"; "AeroSuite.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Begin VB.Form FormLaporan 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9900
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormLaporan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPEngine.XPControl XPEngine 
      Left            =   120
      Top             =   7200
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin AeroSuite.AeroCheckBox CekCustom 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      Align           =   0
      Caption         =   "Custom"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   0
      MousePointer    =   0
      MouseIcon       =   "FormLaporan.frx":030A
      Value           =   0
   End
   Begin AeroSuite.AeroCheckBox cekHanyaHariIni 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      Align           =   0
      Caption         =   "Hanya Hari ini"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   0
      MousePointer    =   0
      MouseIcon       =   "FormLaporan.frx":0326
      Value           =   0
   End
   Begin AeroSuite.AeroCheckBox CekKeseluruhan 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Align           =   0
      Caption         =   "Keseluruhan"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   0
      MousePointer    =   0
      MouseIcon       =   "FormLaporan.frx":0342
      Value           =   0
   End
End
Attribute VB_Name = "FormLaporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    With Me
        .CekKeseluruhan.Value = Checked
        .cekHanyaHariIni.Value = Unchecked
        .CekCustom.Value = Unchecked
    End With
    XPEngine.StartEngine
End Sub

Private Sub CekCustom_Click()
    CekCustom.Value = Checked
    If CekCustom.Value = Checked Then
        CekKeseluruhan.Value = Unchecked
        cekHanyaHariIni.Value = Unchecked
    End If
End Sub

Private Sub cekHanyaHariIni_Click()
    cekHanyaHariIni.Value = Checked
    If cekHanyaHariIni.Value = Checked Then
        CekKeseluruhan.Value = Unchecked
        CekCustom.Value = Unchecked
    End If
End Sub

Private Sub CekKeseluruhan_Click()
    CekKeseluruhan.Value = Checked
    If CekKeseluruhan.Value = Checked Then
        cekHanyaHariIni.Value = Unchecked
        CekCustom.Value = Unchecked
    End If
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub


