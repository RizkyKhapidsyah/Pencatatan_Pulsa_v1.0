Attribute VB_Name = "Engine"
Option Explicit

Public CN As New ADODB.Connection
Public RS As New ADODB.Recordset
Public LI As ListItem
Public Objek As Control
Public X As Integer
Public Y As Integer
Public Z As Integer
Public Kalimat As String
Public LokasiFile As String
Public Pesan As Integer
Public R As Long 'VARIABEL YANG DIPAKAI UNTUK MEMBUKA COMBOBOX TANPA MENGKLIKNYA
Public Result As String

'FUNGSI API YANG DIPAKAI UNTUK MENAMPILKAN ISI COMBOBOX TANPA MENGKLIKNYA DAN MELEBARKAN COMBOBOX
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg _
As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const CB_SHOWDROPDOWN = &H14F
'KONSTANTA YANG DIGUNAKAN UNTUK MELEBARKAN ISI COMBOBOX
Public Const CB_SETDROPPEDWIDTH = &H160



'FUNGSI API YANG DIGUNAKAN UNTUK MENGEKSPOR DATA DARI DATAGRID KE MICROSOFT WORD
Public wrdApp As Word.Application 'MS Word object
Public wrdDoc As Word.Document 'MS Word Document
Public wrdSelection As Word.Selection 'MS Word Selection

Public Declare Function GetSystemMenu Lib "user32" _
(ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, _
ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&
Public Sub RemoveCancelMenuItem(frm As Form)
Dim hSysMenu As Long
  'Ambil menu system untuk form ini
  hSysMenu = GetSystemMenu(frm.hwnd, 0)
  'Hilangkan tombol Close (X)
  Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
  'Hilangkan pemisah yang melalui tombol Close tersebut
  Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
End Sub
'Walaupun tombol "Close" di pojok kanan atas form tidak 'dapat diklik karena sudah disabled, Anda masih bisa 'menutup form dengan menggunakan tombol Alt-F4. Agar 'form juga tidak dapat ditutup dengan menggunakan
'Alt-'F4, Anda harus menahannya di event procedure 'Form_QueryUnload dengan meng-assignment nilai 'parameter Cancel = -1.
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = -1 'Jadi, Alt-F4 juga tidak berfungsi!
End Sub




Public Sub Nyambungg()
If CN.State = adStateOpen Then CN.Close
    CN.CursorLocation = adUseClient
    CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data.rdb;Persist Security Info=False"
End Sub

Public Sub PusatError()
    If Err.Number = 3021 Then
        MsgBox "Silahkan pilih data yang akan dimasukkan ke nama penerima telepon!", vbExclamation + vbOKOnly, ""
    ElseIf Err.Number = -2147467259 Then
        MsgBox "Ada Kesalahan. Nomor Telepon sudah terdaftar!", vbExclamation + vbOKOnly, ""
        FormTambahEditListTelepon.textNomorTelepon.SetFocus
        With FormListTelepon
            .AdodcMain.Refresh
            .AturKontrol
        End With
    Else
        MsgBox Err.Description & vbCrLf & _
                Err.Number, vbCritical + vbOKOnly, "Error"
    End If
End Sub

Public Sub CheckSoftware(X As Form)
On Error GoTo Pesan
    Dim SaveTitle$
    If App.PrevInstance Then
        SaveTitle$ = App.Title
        MsgBox "Program ini sedang dijalankan!", _
               vbCritical, "Sedang Dijalankan"
        App.Title = ""
        X.Caption = ""
        AppActivate SaveTitle$
        SendKeys "%{ENTER}", True
        End
    End If
    Exit Sub
Pesan:
    End
    Exit Sub
End Sub

Public Sub DefaultFormat()
    Select Case FormPengaturan.cmbDefaultSimpan.ListIndex
    Case Is = 0
        FormHasilPencarian.CommonDialog1.FilterIndex = 2
        FormHirarkiView.CommonDialog1.FilterIndex = 2
        FormManage.CommonDialog1.FilterIndex = 2
        FormBaru.CommonDialog1.FilterIndex = 2
    Case Is = 1
        FormHasilPencarian.CommonDialog1.FilterIndex = 3
        FormHirarkiView.CommonDialog1.FilterIndex = 3
        FormManage.CommonDialog1.FilterIndex = 3
        FormBaru.CommonDialog1.FilterIndex = 3
    End Select
End Sub
