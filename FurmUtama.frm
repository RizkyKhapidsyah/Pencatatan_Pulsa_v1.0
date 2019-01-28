VERSION 5.00
Object = "{62E1564C-1E05-44A7-A921-FD8347F324A5}#1.0#0"; "HookMenu.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form FormUtama 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RikySoft Pencatatan Pulsa v1.0"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   15825
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FurmUtama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   15825
   StartUpPosition =   2  'CenterScreen
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   4440
      Top             =   480
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   4
      Bmp:1           =   "FurmUtama.frx":030A
      Key:1           =   "#menuListTelepon"
      Bmp:2           =   "FurmUtama.frx":0732
      Key:2           =   "#menuKalender"
      Bmp:3           =   "FurmUtama.frx":149A
      Key:3           =   "#menuKeluar"
      Bmp:4           =   "FurmUtama.frx":18C2
      Key:4           =   "#menuDataReferensi"
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
   Begin VB.Timer TimerWaktuDanTanggal 
      Interval        =   10
      Left            =   2760
      Top             =   480
   End
   Begin MSComctlLib.StatusBar StatusBawah 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   9
      Top             =   8145
      Width           =   15825
      _ExtentX        =   27914
      _ExtentY        =   794
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin XPEngine.XPControl XPEngine 
      Left            =   2400
      Top             =   600
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin MSComctlLib.ListView LV 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   10821
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin isButton3.isButton cmBaru 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "FurmUtama.frx":1CEA
      Style           =   6
      Caption         =   "&Baru"
      IconAlign       =   1
      iNonThemeStyle  =   0
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
   Begin isButton3.isButton cmManage 
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "FurmUtama.frx":1E44
      Style           =   6
      Caption         =   "&Manage"
      IconAlign       =   1
      iNonThemeStyle  =   0
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
   Begin isButton3.isButton cmKeluar 
      Height          =   495
      Left            =   14400
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "FurmUtama.frx":2296
      Style           =   6
      Caption         =   "&Keluar"
      IconAlign       =   1
      iNonThemeStyle  =   0
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
   Begin isButton3.isButton cmHirarkiView 
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "FurmUtama.frx":23F0
      Style           =   6
      Caption         =   "     &Hirarki View"
      IconAlign       =   1
      iNonThemeStyle  =   0
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
   Begin isButton3.isButton cmRefresh 
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "FurmUtama.frx":254A
      Style           =   6
      Caption         =   "&Refresh"
      IconAlign       =   1
      iNonThemeStyle  =   0
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
   Begin isButton3.isButton cmTentang 
      Height          =   495
      Left            =   12960
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "FurmUtama.frx":26A4
      Style           =   6
      Caption         =   "&Tentang"
      IconAlign       =   1
      iNonThemeStyle  =   0
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
   Begin isButton3.isButton cmPengaturan 
      Height          =   495
      Left            =   7320
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "FurmUtama.frx":27FE
      Style           =   6
      Caption         =   "    &Pengaturan"
      IconAlign       =   1
      iNonThemeStyle  =   0
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
   Begin isButton3.isButton cmProperties 
      Height          =   495
      Left            =   5880
      TabIndex        =   8
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "FurmUtama.frx":4FB0
      Style           =   6
      Caption         =   "&Properties"
      IconAlign       =   1
      iNonThemeStyle  =   0
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
   Begin MSAdodcLib.Adodc AdodcMain 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdodcTools 
      Height          =   330
      Left            =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Menu MenuData 
      Caption         =   "Data"
      Begin VB.Menu menuListTelepon 
         Caption         =   "List Telepon"
      End
      Begin VB.Menu menuDataReferensi 
         Caption         =   "Data Referensi"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuDataRegistrasiDeposit 
         Caption         =   "Data Registrasi Deposit"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu menuKeluar 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu menuTool 
      Caption         =   "Tool"
      Begin VB.Menu menuKalender 
         Caption         =   "Kalender"
      End
   End
   Begin VB.Menu menuView 
      Caption         =   "View"
      Begin VB.Menu menuStatusBawah 
         Caption         =   "Status Bawah"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "FormUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    On Error GoTo HancurkanError
    Nyambungg
    With AdodcMain
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From TablePulsa order by Tahun asc, Bulan, Tanggal"
        .Refresh
    End With
    With LV
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Tanggal", 1400
        .ColumnHeaders.Add , , AdodcMain.Recordset.Fields(4).Name, 1500, vbCenter
        .ColumnHeaders.Add , , AdodcMain.Recordset.Fields(5).Name, 1500, vbCenter
        .ColumnHeaders.Add , , AdodcMain.Recordset.Fields(6).Name, 1500, vbCenter
        .ColumnHeaders.Add , , AdodcMain.Recordset.Fields(7).Name, 1500, vbCenter
        .ColumnHeaders.Add , , AdodcMain.Recordset.Fields(8).Name, 1500, vbCenter
        .ColumnHeaders.Add , , AdodcMain.Recordset.Fields(9).Name, 1500, vbCenter
        .ColumnHeaders.Add , , AdodcMain.Recordset.Fields(10).Name, 1500, vbCenter
        .ColumnHeaders.Add , , AdodcMain.Recordset.Fields(11).Name, 1700, vbCenter
        .ColumnHeaders.Add , , AdodcMain.Recordset.Fields(12).Name, 1300, vbCenter
        .ColumnHeaders.Add , , AdodcMain.Recordset.Fields(13).Name, 1500, vbCenter
        .ColumnHeaders.Add , , AdodcMain.Recordset.Fields(14).Name, 1500, vbCenter
        .View = lvwReport
        .Sorted = True
        If FormPengaturan.cekGridLines.Value = Checked Then .Gridlines = True
    End With
        LV.ListItems.Clear
        Do Until AdodcMain.Recordset.EOF
        Set LI = LV.ListItems.Add(, , AdodcMain.Recordset.Fields(0).Value & ", " & AdodcMain.Recordset.Fields(1).Value & "-" & AdodcMain.Recordset.Fields(2).Value & "-" & AdodcMain.Recordset.Fields(3).Value)
            LI.SubItems(1) = AdodcMain.Recordset.Fields(4).Value
            LI.SubItems(2) = AdodcMain.Recordset.Fields(5).Value
            LI.SubItems(3) = AdodcMain.Recordset.Fields(6).Value
            LI.SubItems(4) = AdodcMain.Recordset.Fields(7).Value
            LI.SubItems(5) = AdodcMain.Recordset.Fields(8).Value
            LI.SubItems(6) = "Rp " & AdodcMain.Recordset.Fields(9).Value
            LI.SubItems(7) = "Rp " & AdodcMain.Recordset.Fields(10).Value
            LI.SubItems(8) = "Rp " & AdodcMain.Recordset.Fields(11).Value
            LI.SubItems(9) = AdodcMain.Recordset.Fields(12).Value
            LI.SubItems(10) = "Rp " & AdodcMain.Recordset.Fields(13).Value
            LI.SubItems(11) = AdodcMain.Recordset.Fields(14).Value
            AdodcMain.Recordset.MoveNext
        Loop
        AdodcMain.Refresh
    With LV
        .ColumnHeaders(1).Width = 1335.118
        .ColumnHeaders(2).Width = 1349.858
        .ColumnHeaders(3).Width = 1500.095
        .ColumnHeaders(4).Width = 915.0237
        .ColumnHeaders(5).Width = 1154.835
        .ColumnHeaders(6).Width = 1500.095
        .ColumnHeaders(7).Width = 1140.095
        .ColumnHeaders(8).Width = 959.8111
        .ColumnHeaders(9).Width = 1100.0945
        .ColumnHeaders(10).Width = 2000.26
        .ColumnHeaders(11).Width = 1065.26
        .ColumnHeaders(12).Width = 1065.26
    End With
    AturStatusBawah
        RemoveCancelMenuItem Me
        Me.Picture = LoadPicture(App.Path & "\image\bannerPencatatanPulsa.jpg")
        If FormPengaturan.CekToolTipText.Value = Checked Then
            LV.ToolTipText = "Penampil data secara umum"
            cmBaru.ToolTipText = "'Klik' untuk menambah data baru!"
            cmManage.ToolTipText = "'Klik' untuk memanage data yang telah tersimpan"
            cmHirarkiView.ToolTipText = "'Klik' untuk menampilkan data secara hirarki"
            cmRefresh.ToolTipText = "'Klik' untuk menyegarkan data"
            cmProperties.ToolTipText = "'Klik' untuk melihat penjelasan data"
            cmPengaturan.ToolTipText = "'klik' untuk mengatur settingan program"
            cmTentang.ToolTipText = "'Klik' untuk melihat tentang aplikasi program"
            cmKeluar.ToolTipText = "'Klik' Untuk keluar dari program"
        ElseIf FormPengaturan.CekToolTipText.Value = Unchecked Then
            LV.ToolTipText = Empty
            cmBaru.ToolTipText = Empty
            cmManage.ToolTipText = Empty
            cmHirarkiView.ToolTipText = Empty
            cmRefresh.ToolTipText = Empty
            cmProperties.ToolTipText = Empty
            cmPengaturan.ToolTipText = Empty
            cmTentang.ToolTipText = Empty
            cmKeluar.ToolTipText = Empty
        End If
    XPEngine.StartEngine
    AturThema
    Call CheckSoftware(Me)
Exit Sub
HancurkanError:
    PusatError
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
Sub AturStatusBawah()
    With StatusBawah
        .Panels.Item(1).Width = 7300
        .Panels.Item(2).Width = 600
        .Panels.Item(3).Width = 600
        .Panels.Item(4).Width = 800
        .Panels.Item(5).Width = 800
        .Panels.Item(6).Width = 2800
        .Panels.Item(7).Width = 1400
        .Panels.Item(8).Width = 1400
        .Panels.Item(1).Alignment = sbrLeft
        .Panels.Item(2).Alignment = sbrCenter
        .Panels.Item(3).Alignment = sbrCenter
        .Panels.Item(4).Alignment = sbrCenter
        .Panels.Item(5).Alignment = sbrCenter
        .Panels.Item(6).Alignment = sbrCenter
        .Panels.Item(7).Alignment = sbrCenter
        .Panels.Item(8).Alignment = sbrCenter
        .Panels.Item(1).ToolTipText = "Status Utama"
        .Panels.Item(2).ToolTipText = "Jumlah Data"
        .Panels.Item(3).ToolTipText = "Jumlah Cell (LV)"
        .Panels.Item(4).ToolTipText = "Jumlah Yang Masih Hutang"
        .Panels.Item(5).ToolTipText = "Jumlah Yang Sudah Lunas"
        .Panels.Item(6).ToolTipText = "Pengisi Terakhir"
        .Panels.Item(7).ToolTipText = "Tanggal Saat Ini"
        .Panels.Item(8).ToolTipText = "Waktu Saat Ini"
        .Panels.Item(1).Text = "Database Ready."
        .Panels.Item(2).Text = AdodcMain.Recordset.RecordCount
        .Panels.Item(3).Text = Val(LV.ColumnHeaders.Count) * Val(LV.ListItems.Count)
            AdodcTools.ConnectionString = CN.ConnectionString
            AdodcTools.RecordSource = "Select * from TablePulsa where Status_Bayar = 'Hutang'"
            AdodcTools.Refresh
        .Panels.Item(4).Text = AdodcTools.Recordset.RecordCount
            AdodcTools.ConnectionString = CN.ConnectionString
            AdodcTools.RecordSource = "Select * from TablePulsa where Status_Bayar = 'Lunas'"
            AdodcTools.Refresh
        .Panels.Item(5).Text = AdodcTools.Recordset.RecordCount
            AdodcMain.Recordset.MoveLast
        .Panels.Item(6).Text = AdodcMain.Recordset.Fields(5).Value & "(" & AdodcMain.Recordset.Fields(4).Value & ") : " & AdodcMain.Recordset.Fields(8).Value
    End With
End Sub


Private Sub cmBaru_Click()
    With FormBaru
        .Caption = "Data Baru"
        .Show vbModal, Me
    End With
End Sub


Private Sub cmHirarkiView_Click()
    If AdodcMain.Recordset.RecordCount = 0 Then
        MsgBox "Hirarki View tidak dapat ditampilkan karena data masih kosong!", vbExclamation + vbOKOnly, ""
    Else
        FormHirarkiView.Show vbModal, Me
    End If
End Sub

Private Sub cmKeluar_Click()
    X = MsgBox("Anda yakin ingin keluar?", vbQuestion + vbYesNo, "Keluar?")
    If X = vbYes Then
        End
    End If
End Sub

Private Sub cmManage_Click()
    If AdodcMain.Recordset.RecordCount = 0 Then
        MsgBox "Manage data tidak dapat ditampilkan karena data masih kosong!", vbExclamation + vbOKOnly, ""
    Else
        FormManage.Show vbModal, Me
    End If
End Sub

Private Sub cmPengaturan_Click()
    FormPengaturan.Show vbModal, Me
End Sub

Private Sub cmProperties_Click()
    MsgBox "Jenis Database : MyISAM" & vbCrLf & _
            "Jumlah Data : " & AdodcMain.Recordset.RecordCount & vbCrLf & _
            "Jumlah Cell (LV) : " & Val(LV.ColumnHeaders.Count) * Val(LV.ListItems.Count) & vbCrLf & _
            "Jumlah Cell (DG): " & Val(AdodcMain.Recordset.RecordCount) * Val(AdodcMain.Recordset.Fields.Count), vbInformation + vbOKOnly, "Properties"
            
End Sub

Public Sub cmRefresh_Click()
    AturKontrol
End Sub

Private Sub cmTentang_Click()
    MsgBox "Pencatatan Pulsa v1.0 by Rizky Khafitsyah." & vbCrLf & _
            "Copyright (c)_2016. All Right Reserved." & vbCrLf & vbCrLf & _
            "Untuk Software lainnya silahkan kunjungi : http://rikymetalist.blogspot.com", vbInformation + vbOKOnly, "Tentang.."
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub


Private Sub menuDataReferensi_Click()
    FormDataReferensi.Show vbModal, Me
End Sub

Private Sub MenuDataRegistrasiDeposit_Click()
    FormDataRegistrasiDeposit.Show vbModal, Me
End Sub

Private Sub menuKalender_Click()
    With FormKalender
        .Caption = "Kalender"
        .cmOK_FormUtama.Visible = True
        .cmOK_UntukDataBaru.Visible = False
        .cmTutup.Visible = False
        .Show vbModal, Me
    End With
End Sub

Private Sub menuKeluar_Click()
    X = MsgBox("Anda yakin ingin keluar?", vbQuestion + vbYesNo, "Keluar?")
    If X = vbYes Then
        End
    End If
End Sub

Private Sub menuListTelepon_Click()
    With FormListTelepon
        .cmMasukan.Enabled = False
        .cmTambah.Enabled = True
        .cmEdit.Enabled = True
        .cmCari.Enabled = True
        .cmHapus.Enabled = True
        .cmRefresh.Enabled = True
        .Show vbModal, Me
    End With
End Sub

Sub AturMenuStatusBawah()
    Select Case menuStatusBawah.Checked
    Case Is = False
        StatusBawah.Visible = True
        Me.Height = 9315
        menuStatusBawah.Checked = True
    Case Is = True
        StatusBawah.Visible = False
        Me.Height = 8899
        menuStatusBawah.Checked = False
    End Select
End Sub

Private Sub menuStatusBawah_Click()
    AturMenuStatusBawah
End Sub

Private Sub TimerWaktuDanTanggal_Timer()
    With StatusBawah
        .Panels.Item(7).Text = Day(Date) & " - " & Month(Date) & " - " & Year(Date)
        .Panels.Item(8).Text = Hour(Time) & ":" & Minute(Time) & ":" & Second(Time)
    End With
End Sub
