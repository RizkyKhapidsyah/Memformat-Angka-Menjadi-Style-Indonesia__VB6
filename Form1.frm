VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memformat Angka Menjadi Style Indonesia"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   2160
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Masukkan bilangan di Text1, lalu tekan Enter untuk
'Menampilkan angka dalam format Indonesia
'Created by Rizky Khapidsyah
'Source code program dimulai dari sini

Private Sub Command1_Click()
  MsgBox FormatAngka(Text1.Text)
  Text1.SetFocus
  SendKeys "{Home}+{End}"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  Command1.Default = True
End Sub

'Parameter Angka bertipe Variant, untuk mengatasi
'apakah input dalam tipe data apapun
Public Function FormatAngka(Angka) As String

'Variabel yang digunakan di fungsi
Dim Jumlah As Integer, Jumlah1 As Integer
Dim i As Integer, j As Integer, k As Integer
Dim strAngka As String, strAngka1 As String
Dim strAngkaFull As String
Dim strTemp As String, strTemp1 As String
   'Tampung nilai angka ke dalam variabel string
    strAngka = CStr(Trim(Angka))
   'Karena parameternya bertipe angka bulat, maka tidak
   'boleh ada karakter lainnya (termasuk titik dan
   'koma) selain karakter angka saja...
   If InStr(1, strAngka, ".") > 0 Or _
      InStr(1, strAngka, ",") > 0 Or _
      Not IsNumeric(Angka) Then
      MsgBox "Harus bilangan bulat dan tidak" & _
       vbCrLf & "boleh mengandung karakter" & _
       vbCrLf & "titik atau koma.", _
             vbCritical, "Bukan Bilangan Bulat"
      Exit Function
   End If
   
   'Tambahkan dua angka nol di belakang string strAngka
   strAngkaFull = strAngka & "00"
      
   'Tampung jumlah digit
   Jumlah = Len(Trim(strAngkaFull))
   
   'Inisialisasi untuk counter menghitung per karakter
   j = 0
   strTemp = ""
   
   'Ulangi setiap karakter mulai dari kanan ke kiri
   For i = Jumlah To 1 Step -1  'Step -1 = berkurang 1
   
      j = j + 1   'Counter untuk semua karakter
      k = k + 1   'Counter untuk letak tanda titik
      
      'Tampung setiap satu karakter ke strTemp
      strTemp = strTemp & Mid(strAngkaFull, i, 1)
      
      'Jika counter = 2 (untuk letak tanda koma
      'desimal)
      If j = 2 Then
         
         'Tambahkan karakter koma
         strTemp = strTemp & ","
         
         'Inisialisasi kembali counter untuk titik
         k = 0
      End If
      
      'Jika counter utk titik = 3 dan
      'belum mencapai digit akhir paling kiri (awal).
      'Hal ini untuk menghindari karakter titik di
      'akhir..
      If k = 3 And i <> 1 Then
         
         'Tambahkan karakter titik
         strTemp = strTemp & "."
         
         'Inisialisasi kembali counter untuk
         'menentukan posisi titik
         k = 0
      End If
   Next i  'Maju ke karakter berikutnya
   
   'Tampung jumlah karakter strTemp yang berasal
   'dari iterasi di atas ini
   Jumlah1 = Len(Trim(strTemp))
   
   'Iterasi berikut untuk membalikkan posisi bilangan
   For i = Jumlah1 To 1 Step -1
      strTemp1 = strTemp1 & Mid(strTemp, i, 1)
   Next i
   
   'Kembalikan nilai string yg fix ke fungsi
     'FormatAngka
   FormatAngka = strTemp1
   
End Function


