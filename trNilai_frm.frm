VERSION 5.00
Begin VB.Form trNilai_frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nilai Guru"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   3735
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Batal"
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   1800
      Width           =   975
   End
   Begin VB.ComboBox cbokategori 
      Height          =   315
      Left            =   1080
      TabIndex        =   6
      Top             =   1320
      Width           =   2535
   End
   Begin VB.ComboBox cboguru 
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nilai"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   300
   End
   Begin VB.Label lblidTrans 
      AutoSize        =   -1  'True
      Caption         =   "IDTrans"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   570
   End
   Begin VB.Label lblidkategori 
      AutoSize        =   -1  'True
      Caption         =   "idperiode"
      Height          =   195
      Left            =   4080
      TabIndex        =   7
      Top             =   1320
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Kriteria"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label lblidguru 
      AutoSize        =   -1  'True
      Caption         =   "idperiode"
      Height          =   195
      Left            =   4080
      TabIndex        =   4
      Top             =   840
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Guru"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   345
   End
   Begin VB.Label lblidperiode 
      AutoSize        =   -1  'True
      Caption         =   "idperiode"
      Height          =   195
      Left            =   4080
      TabIndex        =   1
      Top             =   480
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Periode"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   540
   End
End
Attribute VB_Name = "trNilai_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub isiGuru()
cboguru.Clear
Data1.RecordSource = "msguru"
Data1.Refresh
With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            cboguru.AddItem !namaguru
            .MoveNext
        Loop
        cboguru.ListIndex = 0
    End If
End With
End Sub

Sub isiKategori()
cbokategori.Clear
Data1.RecordSource = "mskategori"
Data1.Refresh
With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            cbokategori.AddItem !kategori
            .MoveNext
        Loop
        cbokategori.ListIndex = 0
    End If
End With
End Sub

Private Sub cboguru_Click()
Data1.RecordSource = "select id from msguru where namaguru='" & cboguru.Text & "'"
Data1.Refresh
Me.lblidguru.Caption = Data1.Recordset!id
End Sub

Private Sub cbokategori_Click()
Data1.RecordSource = "select id from mskategori where kategori='" & cbokategori.Text & "'"
Data1.Refresh
Me.lblidkategori.Caption = Data1.Recordset!id
End Sub

Private Sub Command1_Click()
If Me.Text1 = "" Or Val(Text1) < 1 Or Val(Text1) > 5 Then
    MsgBox "Nilai tidak boleh kosong/kurang dari 1/lebih dari 5", vbInformation, "Validasi"
Else
    If Me.lblidTrans.Caption = "" Then
        Data1.RecordSource = "trnilai"
    Else
        Data1.RecordSource = "select * from trnilai where id=" & Me.lblidTrans.Caption
    End If
    Data1.Refresh
    With Data1.Recordset
    If Me.lblidTrans.Caption = "" Then
        .AddNew
    Else
        .Edit
    End If
    !periode = Me.lblidperiode.Caption
    !guru = Me.lblidguru.Caption
    !kriteria = Me.lblidkategori.Caption
    !nilai = Text1
    .Update
    End With
    If Me.lblidTrans.Caption = "" Then
        Dim tny As String
        tny = MsgBox("Apakah akan tambah lagi?", vbYesNo, "Tambah")
        If tny = vbYes Then
            Text1 = ""
            Text1.SetFocus
        Else
            Unload Me
        End If
    Else
        MsgBox "Berhasil Update data", vbInformation, "Update"
        Unload Me
    End If
    
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
