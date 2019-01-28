VERSION 5.00
Object = "{62E1564C-1E05-44A7-A921-FD8347F324A5}#1.0#0"; "HookMenu.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "xTab.ocx"
Begin VB.MDIForm FORM_UTAMA 
   BackColor       =   &H8000000C&
   Caption         =   "RikySoft - Bisnis Rumahan"
   ClientHeight    =   2940
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4440
   Icon            =   "FORM_UTAMA.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin XPEngine.XPControl XPEngine 
      Left            =   1080
      Top             =   1320
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   2040
      Top             =   2040
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu menuFile 
      Caption         =   "File"
      Begin VB.Menu menuBisnisPulsa 
         Caption         =   "Bisnis Pulsa"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu menuKeluar 
         Caption         =   "Keluar"
      End
   End
End
Attribute VB_Name = "FORM_UTAMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
   Call CheckSoftware(Me)
    WindowState = vbMaximized
    RemoveCancelMenuItem Me
    Me.Picture = LoadPicture(App.Path & "\image\Banner.jpg")
    XPEngine.StartEngine
End Sub

Private Sub menuBisnisPulsa_Click()
    FormUtama.Show vbModal, Me
End Sub

Private Sub menuKeluar_Click()
    x = MsgBox("Anda yakin ingin keluar?", vbQuestion + vbYesNo, "Keluar?")
    If x = vbYes Then
        End
    End If
End Sub
