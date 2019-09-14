VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form IMPORTEXCELDAT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IMPORT  FILES"
   ClientHeight    =   2325
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "PROSES"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   5415
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Text            =   "10"
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "2"
      Top             =   1200
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "DARI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "SAMPAI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "IMPORTEXCELDAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error GoTo ERRNUM
With CommonDialog1
        .DialogTitle = "Pilih File Excelnya (.xls)"
        .InitDir = App.Path
        .Filter = "SQL Files (*.xls)|*.xls"
        'jika menggunakan file excel 2007 keatas
        'untuk .Filter = "SQL Files (*.xls)|*.xls" '
        'Ganti dengan .Filter = "SQL Files (*.xlsx)|*.xlsx"
        .ShowOpen
    End With
    'menampilkan nama filenya di textbox
    Text1.Text = CommonDialog1.filename
ERRNUM:
If Err.Number <> 0 Then MsgBox Err.Description & vbCrLf & Err.Number
End Sub

Private Sub StartExcel()
On Error GoTo Err:
'Pertama, ambil object Excel, dan jika error
'lompat ke label Err di paling bawah Sub ini,
'dan buat object Excel. Error terjadi jika
'object Excel belum dibuat
Set MyExcel = New Excel.Application
'Set Excel = GetObject(, "Excel.Application")
Exit Sub
Err:
'Buat object Excel jika belum ada.
MsgBox Err.Description
End Sub

Private Sub CloseWorkSheet()
On Error Resume Next
'Tutup Excel workbook
MyExcelWBk.Close
'Keluar dari aplikasi Excel
MyExcel.Quit
End Sub

Private Sub ClearExcelMemory()
On Error Resume Next
'Sebelum membersihkan memory, periksa terlebih dulu
'object yang akan dibersihkan...
Set MyExcelWS = Nothing
Set MyExcelWBk = Nothing
Set MyExcel = Nothing
End Sub





Private Sub Command2_Click()
On Error GoTo ERRNUM

StartExcel
Set ExcelWBk = Excel.Workbooks.Open(CommonDialog1.filename)
'Print "Berhasil membuka file ..."
'Akses Sheet pertama (1)
'Jika ingin mengakses sheet kedua, ganti
'(1) menjadi (2), dst...
Set MyExcelWS = ExcelWBk.Worksheets(1)
'Tampilkan status di form
'Print "Berhasil membaca Sheet1 ..."
'Lakukan proses di worksheet ExcelWS
With MyExcelWS
Dim i, j As Integer
i = Val(Text2.Text)
j = Val(Text3.Text)
Dim strData As String
'MsgBox .Rows.Count
'MsgBox .Columns.Count
'Baca mulai dari baris pertama sampai kelima
For i = i To j
'Tampung ke sebuah variabel
'strData = strData & .Cells(i, 1) & vbCrLf
rsTemp17.AddNew
rsTemp17.Fields(0).Value = .Cells(i, 1)
rsTemp17.Fields(1).Value = .Cells(i, 2)
rsTemp17.Fields(2).Value = .Cells(i, 3)
rsTemp17.Fields(3).Value = .Cells(i, 4)
rsTemp17.Fields(4).Value = .Cells(i, 5)
rsTemp17.Fields(5).Value = .Cells(i, 6)
rsTemp17.Fields(6).Value = .Cells(i, 7)
rsTemp17.Fields(7).Value = .Cells(i, 8)
rsTemp17.Fields(8).Value = .Cells(i, 9)
rsTemp17.Update
Next i
End With
'Tampilkan ke layar
'MsgBox strData
'Setelah selesai, jangan lupa tutup worksheet
CloseWorkSheet
'Tampilkan status di form
'Print "Berhasil menutup worksheet dan file Excel ..."
'Jangan lupa pula, selalu bersihkan memory yang
'digunakan oleh object Excel
ClearExcelMemory
'Tampilkan status di form
'Print "Berhasil membersihkan memory Excel ..."
'Tampilkan pesan
MsgBox "Import File Selesai!", vbInformation, "Mantap"
Me.Cls
Unload Me
ERRNUM:
If Err.Number <> 0 Then MsgBox Err.Description & vbCrLf & Err.Number

End Sub

