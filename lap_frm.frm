VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form lap_frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9900
   Begin VB.ListBox List3 
      Height          =   3375
      Left            =   7920
      TabIndex        =   8
      Top             =   3120
      Width           =   1815
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   120
      TabIndex        =   7
      Top             =   5040
      Width           =   7695
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   7695
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   7200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Proses Topsis"
      Height          =   495
      Left            =   8760
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   240
      Width           =   2295
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "lap_frm.frx":0000
      Height          =   2295
      Left            =   120
      OleObjectBlob   =   "lap_frm.frx":0014
      TabIndex        =   3
      Top             =   720
      Width           =   9615
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   5655
   End
   Begin VB.Label lblidperiode 
      Caption         =   "Label2"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Periode"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "lap_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Data2.RecordSource = "select id from msperiode where periode='" & Combo1.Text & "'"
Data2.Refresh
If Data2.Recordset.BOF Then
    lblidperiode.Caption = ""
Else
    Me.lblidperiode.Caption = Data2.Recordset!Id
    isiGrid
End If
End Sub

Private Sub Command1_Click()
Label2.Caption = ""
Data1.Refresh
Dim x1, x2, x3, x4, x5 As Single
x1 = 0
x2 = 0
x3 = 0
x4 = 0
x5 = 0
With Data1.Recordset
    .MoveFirst
    Do While Not .EOF
        x1 = x1 + (!loyalitas * !loyalitas)
        x2 = x2 + (!teladan * !teladan)
        x3 = x3 + (!kehadiran * !kehadiran)
        x4 = x4 + (!administrasi * !administrasi)
        x5 = x5 + (!supervisi * !supervisi)
        .MoveNext
    Loop
    Label2.Caption = "X1=" & x1 & ", X2=" & x2 & ", X3=" & x3 & ", X4=" & x4 & ", X5=" & x5
    .MoveFirst
    Dim r As Integer
    r = 1
    Me.List1.Clear
    Do While Not .EOF
        List1.AddItem "R1" & r & vbTab & Format(!loyalitas / x1, "#0.0000") & vbTab & "R2" & r & vbTab & Format(!teladan / x2, "#0.0000") & vbTab & "R3" & r & vbTab & Format(!kehadiran / x3, "#0.0000") & vbTab & "R4" & r & vbTab & Format(!administrasi / x4, "#0.0000") & vbTab & "R5" & r & vbTab & Format(!supervisi / x5, "#0.0000")
        r = r + 1
        .MoveNext
    Loop
    
    Me.List2.Clear
    Dim y, y1, y2, y3, y4, y5 As Double
    Dim maxy1, maxy2, maxy3, maxy4, maxy5 As Double
    Dim miny1, miny2, miny3, miny4, miny5 As Double
    Dim y11, y12, y13, y14, y15 As Double
    Dim y21, y22, y23, y24, y25 As Double
    Dim y31, y32, y33, y34, y35 As Double
    Dim y41, y42, y43, y44, y45 As Double
    Dim y51, y52, y53, y54, y55 As Double
    
    Data2.RecordSource = "mskategori"
    Data2.Refresh
    Data2.Recordset.MoveFirst
    y = 1
    Dim str1, str2, str3, str4, str5 As String
    'loyalitas
    .MoveFirst
    maxy1 = 0
    miny1 = Data2.Recordset!bobot * !loyalitas / x1
    Do While Not .EOF
        y1 = Data2.Recordset!bobot * !loyalitas / x1
        If y1 > maxy1 Then
            maxy1 = y1
        End If
        If y1 < miny1 Then
            miny1 = y1
        End If
        Select Case y
            Case 1
                y11 = y1
            Case 2
                y12 = y1
            Case 3
                y13 = y1
            Case 4
                y14 = y1
            Case 5
                y15 = y1
        End Select
        y = y + 1
        str1 = str1 & vbTab & Format(y1, "#0.0000")
        .MoveNext
    Loop
    'teladan
    y = 1
    Data2.Recordset.MoveNext
    .MoveFirst
    maxy2 = 0
    miny2 = Data2.Recordset!bobot * !teladan / x2
    Do While Not .EOF
        y2 = Data2.Recordset!bobot * !teladan / x2
        If y2 > maxy2 Then
            maxy2 = y2
        End If
        If y2 < miny2 Then
            miny2 = y2
        End If
        Select Case y
            Case 1
                y21 = y2
            Case 2
                y22 = y2
            Case 3
                y23 = y2
            Case 4
                y24 = y2
            Case 5
                y25 = y2
        End Select
        y = y + 1
        str2 = str2 & vbTab & Format(y2, "#0.0000")
        .MoveNext
    Loop
    'kehadiran
    y = 1
    Data2.Recordset.MoveNext
    .MoveFirst
    maxy3 = 0
    miny3 = Data2.Recordset!bobot * !kehadiran / x3
    Do While Not .EOF
        y3 = Data2.Recordset!bobot * !kehadiran / x3
        If y3 > maxy3 Then
            maxy3 = y3
        End If
        If y3 < miny3 Then
            miny3 = y3
        End If
        Select Case y
            Case 1
                y31 = y3
            Case 2
                y32 = y3
            Case 3
                y33 = y3
            Case 4
                y34 = y3
            Case 5
                y35 = y3
        End Select
        y = y + 1
        str3 = str3 & vbTab & Format(y3, "#0.0000")
        .MoveNext
    Loop
    'administrasi
    y = 1
    Data2.Recordset.MoveNext
    .MoveFirst
    maxy4 = 0
    miny4 = Data2.Recordset!bobot * !administrasi / x4
    Do While Not .EOF
        y4 = Data2.Recordset!bobot * !administrasi / x4
        If y4 > maxy4 Then
            maxy4 = y4
        End If
        If y4 < miny4 Then
            miny4 = y4
        End If
        Select Case y
            Case 1
                y41 = y4
            Case 2
                y42 = y4
            Case 3
                y43 = y4
            Case 4
                y44 = y4
            Case 5
                y45 = y4
        End Select
        y = y + 1
        str4 = str4 & vbTab & Format(y4, "#0.0000")
        .MoveNext
    Loop
    'supervisi
    y = 1
    Data2.Recordset.MoveNext
    .MoveFirst
    maxy5 = 0
    miny5 = Data2.Recordset!bobot * !supervisi / x5
    Do While Not .EOF
        y5 = Data2.Recordset!bobot * !supervisi / x5
        If y5 > maxy5 Then
            maxy5 = y5
        End If
        If y5 < miny5 Then
            miny5 = y5
        End If
        Select Case y
            Case 1
                y51 = y5
            Case 2
                y52 = y5
            Case 3
                y53 = y5
            Case 4
                y54 = y5
            Case 5
                y55 = y5
        End Select
        y = y + 1
        str5 = str5 & vbTab & Format(y5, "#0.0000")
        .MoveNext
    Loop
    
    List2.AddItem "Y1" & str1 & vbTab & " Max : " & Format(maxy1, "#0.0000") & vbTab & " Min : " & Format(miny1, "#0.0000")
    List2.AddItem "Y2" & str2 & vbTab & " Max : " & Format(maxy2, "#0.0000") & vbTab & " Min : " & Format(miny2, "#0.0000")
    List2.AddItem "Y3" & str3 & vbTab & " Max : " & Format(maxy3, "#0.0000") & vbTab & " Min : " & Format(miny3, "#0.0000")
    List2.AddItem "Y4" & str4 & vbTab & " Max : " & Format(maxy4, "#0.0000") & vbTab & " Min : " & Format(miny4, "#0.0000")
    List2.AddItem "Y5" & str5 & vbTab & " Max : " & Format(maxy5, "#0.0000") & vbTab & " Min : " & Format(miny5, "#0.0000")
    
    List3.Clear
    Dim d1max, d2max, d3max, d4max, d5max As Double
    d1max = ((maxy1 - y11) ^ 2) + ((maxy2 - y21) ^ 2) + ((maxy3 - y31) ^ 2) + ((maxy4 - y41) ^ 2) + ((maxy5 - y51) ^ 2)
    d2max = ((maxy1 - y12) ^ 2) + ((maxy2 - y22) ^ 2) + ((maxy3 - y32) ^ 2) + ((maxy4 - y42) ^ 2) + ((maxy5 - y52) ^ 2)
    d3max = ((maxy1 - y13) ^ 2) + ((maxy2 - y23) ^ 2) + ((maxy3 - y33) ^ 2) + ((maxy4 - y43) ^ 2) + ((maxy5 - y53) ^ 2)
    d4max = ((maxy1 - y14) ^ 2) + ((maxy2 - y24) ^ 2) + ((maxy3 - y34) ^ 2) + ((maxy4 - y44) ^ 2) + ((maxy5 - y54) ^ 2)
    d5max = ((maxy1 - y15) ^ 2) + ((maxy2 - y25) ^ 2) + ((maxy3 - y35) ^ 2) + ((maxy4 - y45) ^ 2) + ((maxy5 - y55) ^ 2)
    List3.AddItem "D1+" & vbTab & Format(d1max, "#0.0000")
    List3.AddItem "D2+" & vbTab & Format(d2max, "#0.0000")
    List3.AddItem "D3+" & vbTab & Format(d3max, "#0.0000")
    List3.AddItem "D4+" & vbTab & Format(d4max, "#0.0000")
    List3.AddItem "D5+" & vbTab & Format(d5max, "#0.0000")
    
    List3.AddItem ""
    Dim d1min, d2min, d3min, d4min, d5min As Double
    d1min = ((miny1 - y11) ^ 2) + ((miny2 - y21) ^ 2) + ((miny3 - y31) ^ 2) + ((miny4 - y41) ^ 2) + ((miny5 - y51) ^ 2)
    d2min = ((miny1 - y12) ^ 2) + ((miny2 - y22) ^ 2) + ((miny3 - y32) ^ 2) + ((miny4 - y42) ^ 2) + ((miny5 - y52) ^ 2)
    d3min = ((miny1 - y13) ^ 2) + ((miny2 - y23) ^ 2) + ((miny3 - y33) ^ 2) + ((miny4 - y43) ^ 2) + ((miny5 - y53) ^ 2)
    d4min = ((miny1 - y14) ^ 2) + ((miny2 - y24) ^ 2) + ((miny3 - y34) ^ 2) + ((miny4 - y44) ^ 2) + ((miny5 - y54) ^ 2)
    d5min = ((miny1 - y15) ^ 2) + ((miny2 - y25) ^ 2) + ((miny3 - y35) ^ 2) + ((miny4 - y45) ^ 2) + ((miny5 - y55) ^ 2)
    List3.AddItem "D1-" & vbTab & Format(d1min, "#0.0000")
    List3.AddItem "D2-" & vbTab & Format(d2min, "#0.0000")
    List3.AddItem "D3-" & vbTab & Format(d3min, "#0.0000")
    List3.AddItem "D4-" & vbTab & Format(d4min, "#0.0000")
    List3.AddItem "D5-" & vbTab & Format(d5min, "#0.0000")
End With
With result_frm
.Data1.DatabaseName = App.Path & "\dbnilai.mdb"
.Data1.RecordSource = "tempranking"
.Data1.Refresh
If .Data1.Recordset.RecordCount > 0 Then
    .Data1.Recordset.MoveFirst
    Do While Not .Data1.Recordset.EOF
        .Data1.Recordset.Delete
        .Data1.Recordset.MoveNext
    Loop
End If
Data1.Recordset.MoveFirst
Dim v1, v2, v3, v4, v5 As Double
v1 = d1min / (d1min + d1max)
v2 = d2min / (d2min + d2max)
v3 = d3min / (d3min + d3max)
v4 = d4min / (d4min + d4max)
v5 = d5min / (d5min + d5max)
Do While Not Data1.Recordset.EOF
    .Data1.Recordset.AddNew
    .Data1.Recordset!namaguru = Data1.Recordset!namaguru
    .Data1.Recordset.Update
    Data1.Recordset.MoveNext
Loop
.Data1.Recordset.MoveFirst
.Data1.Recordset.Edit
.Data1.Recordset!topsis = Format(v1, "#0.0000")
.Data1.Recordset.Update

.Data1.Recordset.MoveNext
.Data1.Recordset.Edit
.Data1.Recordset!topsis = Format(v2, "#0.0000")
.Data1.Recordset.Update

.Data1.Recordset.MoveNext
.Data1.Recordset.Edit
.Data1.Recordset!topsis = Format(v3, "#0.0000")
.Data1.Recordset.Update

.Data1.Recordset.MoveNext
.Data1.Recordset.Edit
.Data1.Recordset!topsis = Format(v4, "#0.0000")
.Data1.Recordset.Update

.Data1.Recordset.MoveNext
.Data1.Recordset.Edit
.Data1.Recordset!topsis = Format(v5, "#0.0000")
.Data1.Recordset.Update

.Data1.RecordSource = "select * from tempranking order by topsis desc"
.Data1.Refresh
.Top = 0
.Left = 0
.Show
End With
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\dbnilai.mdb"
Data1.RecordSource = "trnilai"
Data1.Refresh
Data2.DatabaseName = App.Path & "\dbnilai.mdb"
Data2.RecordSource = "msperiode"
Data2.Refresh
isiCbo
End Sub

Sub isiCbo()
Combo1.Clear
With Data2.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            Combo1.AddItem !periode
            .MoveNext
        Loop
        Combo1.ListIndex = 0
    End If
End With
End Sub

Sub isiGrid()
    Dim sql As String
    'sql = "select trnilai.id,msperiode.periode,msguru.namaguru,mskategori.kategori,nilai from trnilai,msperiode,msguru,mskategori where trnilai.periode=" & Me.lblidperiode.Caption & " and trnilai.periode=msperiode.id and trnilai.guru=msguru.id and trnilai.kriteria=mskategori.id order by trnilai.id"
    sql = "select * from pivot_nilai where period='" & Me.Combo1.Text & "'"
    Data1.RecordSource = sql
    Data1.Refresh
End Sub
