VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00CDEBEB&
   Caption         =   "Skrá hund"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   LinkTopic       =   "Form3"
   ScaleHeight     =   7425
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "Tvískráning"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   960
      Width           =   1095
   End
   Begin VB.CheckBox chkFramraektun 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Framræktun"
      Height          =   255
      Left            =   1320
      TabIndex        =   44
      Top             =   6960
      Width           =   2415
   End
   Begin VB.TextBox txtTitill 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtEigandi 
      Height          =   285
      Left            =   1320
      TabIndex        =   19
      Top             =   5160
      Width           =   3495
   End
   Begin VB.TextBox txtRaektandi 
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      Top             =   5520
      Width           =   3495
   End
   Begin VB.CommandButton cmdRaektandi 
      BackColor       =   &H00AFCDCD&
      Caption         =   "..."
      Height          =   255
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdEigandi 
      BackColor       =   &H00AFCDCD&
      Caption         =   "..."
      Height          =   255
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtOrmerki 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtLitur 
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   3360
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "Eyða færslu"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   480
      Width           =   1095
   End
   Begin VB.CheckBox chkLatinn 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Látinn"
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CheckBox chkGeldur 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Ekki til ræktunar"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00AFCDCD&
      Caption         =   "..."
      Height          =   255
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00AFCDCD&
      Caption         =   "..."
      Height          =   255
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4800
      Width           =   375
   End
   Begin VB.ComboBox cboModir 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4080
      Width           =   3495
   End
   Begin VB.ComboBox cboFadir 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3720
      Width           =   3495
   End
   Begin VB.TextBox txtInnkallsnafn 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtFdags 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtAettbokNr 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmdLoka 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "Loka"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox txtNafn 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.TextBox txtMynd1 
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Top             =   4800
      Width           =   2295
   End
   Begin VB.ComboBox cboLand 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3000
      Width           =   3495
   End
   Begin VB.TextBox txtMynd2 
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Top             =   4440
      Width           =   2295
   End
   Begin VB.ComboBox cboTegund 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CheckBox chkEiginHundur 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Eigin hundur"
      Height          =   255
      Left            =   1320
      TabIndex        =   23
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CheckBox chkMinRaektun 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Mín ræktun"
      Height          =   255
      Left            =   1320
      TabIndex        =   24
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CheckBox chkKlubbhundur 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Deildarhundur"
      Height          =   255
      Left            =   1320
      TabIndex        =   22
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton cmdVistaHund 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "Vista"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6360
      Width           =   1095
   End
   Begin VB.ComboBox cboKyn 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Eigandi:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Ræktandi:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Örmerki:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Litur:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Móðir:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Faðir:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Innkallsnafn:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Kyn:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Fæðingardagur:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Ættbókarnúmer:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Nafn:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Titill:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Fæðingarland:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Mynd:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Andlitsmynd:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Tegund:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   5880
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_db As Database
Private m_nr As Long
Private m_hladid As Boolean
Private pStillingar As CStillingar
Private m_kyn As String

Public Sub Grunnur(fdb As Database)
  Set m_db = fdb
End Sub

Public Sub Stillingar(fStillingar As CStillingar)
  Set pStillingar = fStillingar
End Sub

Public Sub hundur(fnr As Long, Optional fkyn As String)
  m_nr = fnr
  If m_nr > 0 Then Command3.Visible = True
  If m_nr > 0 Then Command4.Visible = True
  m_kyn = fkyn
End Sub

Public Function faHund() As Long
  faHund = m_nr
End Function

Private Sub cmdEigandi_Click()
  Load Form10
  Form10.Grunnur m_db
  Form10.Eigandi
  Form10.EigandiID = txtEigandi.Tag
  Form10.Show vbModal
  If Form10.Nafn <> "" Then
    txtEigandi.Text = Form10.Nafn
    txtEigandi.Tag = Form10.EigandiID
  End If
  Unload Form10
End Sub

Private Sub cmdRaektandi_Click()
  Load Form10
  Form10.Grunnur m_db
  Form10.Raektandi
  Form10.EigandiID = txtRaektandi.Tag
  Form10.Show vbModal
  If Form10.Nafn <> "" Then
    txtRaektandi.Text = Form10.Nafn
    txtRaektandi.Tag = Form10.EigandiID
  End If
  Unload Form10
End Sub

Private Sub Command1_Click()
  Load Form7
  Form7.Grunnur m_db
  Form7.Stillingar pStillingar
  Form7.Show vbModal
  If Form7.SkrarNafn <> "" Then
    txtMynd1.Text = Form7.SkrarNafn
  End If
  Unload Form7
End Sub

Private Sub Command2_Click()
  Load Form7
  Form7.Grunnur m_db
  Form7.Stillingar pStillingar
  Form7.Show vbModal
  If Form7.SkrarNafn <> "" Then
    txtMynd2.Text = Form7.SkrarNafn
  End If
  Unload Form7
End Sub

Private Sub Command3_Click()
  Dim result
  Dim rs As Recordset
  If m_nr <> 0 Then
    Set rs = m_db.OpenRecordset("select * from hundalisti where nrf=" & m_nr & " or nrm=" & m_nr)
    If rs.RecordCount > 0 Then
      MsgBox "Ekki hægt að eyða hundi, er með skráð afkvæmi", vbOKOnly, "Eyða hundi"
      rs.Close
      Exit Sub
    End If
    rs.Close
    
    result = MsgBox("Viltu örugglega eyða út hundinum, " & txtNafn.Text & "?", vbOKCancel, "Eyða færslu.")
    If result = 1 Then
      Set rs = m_db.OpenRecordset("select * from hundar where nr=" & m_nr)
      rs.Delete
      rs.Close
      MsgBox "Færslunni var eytt", , "Eyða hundi"
    End If
  End If
  Me.Hide
End Sub

Private Sub Command4_Click()
  Load Form13
  Form13.Grunnur m_db
  Form13.Stillingar pStillingar
  Form13.nr = m_nr
  Form13.Nafn (txtNafn.Text)
  Form13.Show vbModal
  If Form13.nr <> m_nr Then
    m_nr = Form13.nr
    m_hladid = False
    Form_Load
    Form_activate
  End If
  Unload Form13
End Sub

Private Sub Form_Load()
  m_hladid = False
  Command3.Visible = False
  Command4.Visible = False
End Sub

Private Sub Form_activate()
  If m_hladid = True Then Exit Sub
  m_hladid = True
  Dim rs As Recordset
  cboKyn.Clear
  cboKyn.AddItem "Karlhundur"
  cboKyn.AddItem "Tík"
  cboKyn.AddItem "Ekki vitað"
  cboLand.Clear
  Set rs = m_db.OpenRecordset("select * from land order by heiti")
  While Not rs.EOF
    cboLand.AddItem rs!heiti
    cboLand.ItemData(cboLand.ListCount - 1) = rs!id
    rs.MoveNext
  Wend
  rs.Close
  cboTegund.Clear
  Set rs = m_db.OpenRecordset("select * from tegundir where fciflokkur='" & pStillingar.Tegundaflokkur & "'")
  While Not rs.EOF
    cboTegund.AddItem rs!heiti
    cboTegund.ItemData(cboTegund.ListCount - 1) = rs!id
    rs.MoveNext
  Wend
  rs.Close
  
  cboFadir.Clear
  cboFadir.AddItem ""
  cboFadir.ItemData(cboFadir.ListCount - 1) = 0
  Set rs = m_db.OpenRecordset("select * from hundalisti where kyn = 'kk' and tegundaflokkur='" & pStillingar.Tegundaflokkur & "' order by nafn,fdags")
  While Not rs.EOF
    cboFadir.AddItem rs.Fields("nafn") & " - " & rs.Fields("FDags") 'IIf(rs.Fields("titill") <> "", rs.Fields("titill") & " ", "") &
    cboFadir.ItemData(cboFadir.ListCount - 1) = rs!nr
    rs.MoveNext
  Wend
  rs.Close
  cboModir.Clear
  cboModir.AddItem ""
  cboModir.ItemData(cboModir.ListCount - 1) = 0
  Set rs = m_db.OpenRecordset("select * from hundalisti where kyn = 'kvk' and tegundaflokkur='" & pStillingar.Tegundaflokkur & "' order by nafn,fdags")
  While Not rs.EOF
    cboModir.AddItem rs.Fields("nafn") & " - " & rs.Fields("FDags") 'IIf(rs.Fields("titill") <> "", rs.Fields("titill") & " ", "") &
    cboModir.ItemData(cboModir.ListCount - 1) = rs!nr
    rs.MoveNext
  Wend
  rs.Close
  Birta
End Sub
Private Sub cmdVistaHund_Click()
  MousePointer = 11
  Vista
  MousePointer = 0
  cmdLoka_Click
End Sub

Private Sub cmdLoka_Click()
  Me.Hide
End Sub

Private Sub Birta()
  Dim i As Long
  If m_nr = 0 Then
    txtTitill.Text = ""
    txtNafn.Text = ""
    cboFadir.ListIndex = -1
    cboModir.ListIndex = -1
    txtFdags.Text = ""
    txtAettbokNr.Text = ""
    txtOrmerki.Text = ""
    txtInnkallsnafn.Text = ""
    txtMynd1.Text = ""
    txtMynd2.Text = ""
    txtEigandi.Text = ""
    txtEigandi.Tag = 0
    txtRaektandi.Text = ""
    txtRaektandi.Tag = 0
    txtLitur.Text = ""
    chkLatinn.Value = 0
    chkGeldur.Value = 0
    
    If m_kyn = "kk" Then
      cboKyn.ListIndex = 0
    ElseIf m_kyn = "kvk" Then
      cboKyn.ListIndex = 1
    Else
      cboKyn.ListIndex = -1
    End If
    
    'cboKyn.ListIndex = -1
    cboLand.ListIndex = -1
    For i = 0 To cboTegund.ListCount - 1
      If pStillingar.Tegund = cboTegund.ItemData(i) Then cboTegund.ListIndex = i
    Next
    
    chkEiginHundur.Value = 0
    chkKlubbhundur.Value = 0
    chkMinRaektun.Value = 0
    chkFramraektun.Value = 0
  Else
    Dim rs As Recordset
    Set rs = m_db.OpenRecordset("select * from hundalisti where nr=" & m_nr)
    txtTitill.Text = "" & rs.Fields("Titill")
    txtNafn.Text = "" & rs.Fields("Nafn")
    txtFdags.Text = "" & rs.Fields("Fdags")
    txtAettbokNr.Text = "" & rs.Fields("AettbokarNr")
    txtOrmerki.Text = "" & rs.Fields("ormerki")
    txtLitur.Text = "" & rs.Fields("litur")
    txtInnkallsnafn.Text = "" & rs.Fields("Innkallsnafn")
    txtMynd1.Text = "" & rs.Fields("Mynd1")
    txtMynd2.Text = "" & rs.Fields("Mynd2")
    txtEigandi.Text = "" & rs.Fields("EigandiNafn")
    txtEigandi.Tag = 0 & rs.Fields("EigandiID")
    txtRaektandi.Text = "" & rs.Fields("raektandiNafn")
    txtRaektandi.Tag = 0 & rs.Fields("raektandiID")
    
    chkLatinn.Value = IIf(rs.Fields("latinn") = "já", 1, 0)
    chkGeldur.Value = IIf(rs.Fields("geldur") = "já", 1, 0)
    
    If rs.Fields("kyn") = "kk" Then
      cboKyn.ListIndex = 0
    ElseIf rs.Fields("kyn") = "kvk" Then
      cboKyn.ListIndex = 1
    Else
      cboKyn.ListIndex = 2
    End If
    
    For i = 0 To cboFadir.ListCount - 1
      If rs.Fields("fadirnr") = cboFadir.ItemData(i) Then cboFadir.ListIndex = i
    Next
    
    For i = 0 To cboModir.ListCount - 1
      If rs.Fields("modirnr") = cboModir.ItemData(i) Then cboModir.ListIndex = i
    Next
    
    For i = 0 To cboLand.ListCount - 1
      If rs.Fields("landid") = cboLand.ItemData(i) Then cboLand.ListIndex = i
    Next
    
    For i = 0 To cboTegund.ListCount - 1
      If rs.Fields("tegundid") = cboTegund.ItemData(i) Then cboTegund.ListIndex = i
    Next
    
    chkEiginHundur.Value = IIf(rs.Fields("eiginhundar") = "já", 1, 0)
    chkKlubbhundur.Value = IIf(rs.Fields("klubbhundar") = "já", 1, 0)
    chkMinRaektun.Value = IIf(rs.Fields("minraektun") = "já", 1, 0)
    chkFramraektun.Value = IIf(rs.Fields("latinn") = "fram", 1, 0)
    rs.Close
  End If
End Sub

Private Sub Vista()
  Dim rs As Recordset
  
  If m_nr = 0 Then
    Dim max As Long
    Set rs = m_db.OpenRecordset("select max(nr) as maxnr from hundar")
    If IsNull(rs.Fields("maxnr")) Then
      max = 1
    Else
      max = rs.Fields("maxnr")
    End If
    rs.Close
    Set rs = m_db.OpenRecordset("hundar")
    rs.AddNew
    m_nr = max + 1
    rs.Fields("nr") = m_nr
  Else
    Set rs = m_db.OpenRecordset("select * from hundar where nr=" & m_nr)
    rs.Edit
  End If
    
  rs.Fields("Titill") = txtTitill.Text
  rs.Fields("Nafn") = txtNafn.Text
  rs.Fields("Fdags") = IIf(txtFdags.Text <> "", txtFdags.Text, Null)
  rs.Fields("AettbokarNr") = txtAettbokNr.Text
  rs.Fields("ormerki") = txtOrmerki.Text
  rs.Fields("litur") = txtLitur.Text
  rs.Fields("Innkallsnafn") = txtInnkallsnafn.Text
  rs.Fields("Mynd1") = txtMynd1.Text
  rs.Fields("Mynd2") = txtMynd2.Text
  
  'rs.Fields("eigandiID") = txtEigandi.Tag
  'rs.Fields("raektandiID") = txtRaektandi.Tag
  
  rs.Fields("eigandi") = txtEigandi.Text
  rs.Fields("raektandi") = txtRaektandi.Text
  
  rs.Fields("latinn") = IIf(chkLatinn.Value = 1, "já", IIf(chkFramraektun.Value = 1, "fram", ""))
  rs.Fields("geldur") = IIf(chkGeldur.Value = 1, "já", "")
  
  If cboKyn.ListIndex = 0 Then
    rs.Fields("kyn") = "kk"
  ElseIf cboKyn.ListIndex = 1 Then
    rs.Fields("kyn") = "kvk"
  Else
    rs.Fields("kyn") = "-"
  End If
    
  If cboFadir.ListIndex = -1 Then
    rs.Fields("fadirnr") = 0
  Else
    rs.Fields("fadirnr") = cboFadir.ItemData(cboFadir.ListIndex)
  End If
  
  If cboModir.ListIndex = -1 Then
    rs.Fields("modirnr") = 0
  Else
    rs.Fields("modirnr") = cboModir.ItemData(cboModir.ListIndex)
  End If
  If cboLand.ListIndex <> -1 Then
    rs.Fields("landid") = cboLand.ItemData(cboLand.ListIndex)
  Else
    rs.Fields("landid") = Null
  End If
  rs.Fields("tegundid") = cboTegund.ItemData(cboTegund.ListIndex)
  
  rs.Fields("eiginhundar") = IIf(chkEiginHundur.Value = 1, "já", "-")
  rs.Fields("klubbhundar") = IIf(chkKlubbhundur.Value = 1, "já", "-")
  rs.Fields("minraektun") = IIf(chkMinRaektun.Value = 1, "já", "-")
  
  rs.Update
  rs.Close
End Sub

Private Sub txtFdags_LostFocus()
  Dim rsFd As Recordset
  Dim sFd As String
  Dim result As Integer
  Dim i As Integer
  If txtFdags <> "" And cboFadir.ListIndex = -1 And cboModir.ListIndex = -1 Then
    sFd = Replace(txtFdags.Text, ".", "/")
    Set rsFd = m_db.OpenRecordset("select * from hundalisti where fdags=#" & sFd & "# and nr<>" & m_nr)
    Do While Not rsFd.EOF
      result = MsgBox("Er " & txtNafn.Text & " úr sama goti og " & rsFd!Nafn & vbCrLf & "Foreldrar: " & rsFd!nf & " og " & vbCrLf & rsFd!nm & "?", vbYesNo, "Sama got?")
      If result = 6 Then
    
        For i = 0 To cboFadir.ListCount - 1
          If rsFd.Fields("fadirnr") = cboFadir.ItemData(i) Then cboFadir.ListIndex = i
        Next
        For i = 0 To cboModir.ListCount - 1
          If rsFd.Fields("modirnr") = cboModir.ItemData(i) Then cboModir.ListIndex = i
        Next
        For i = 0 To cboLand.ListCount - 1
          If rsFd.Fields("landid") = cboLand.ItemData(i) Then cboLand.ListIndex = i
        Next
        
        Exit Do
      End If
      rsFd.MoveNext
    Loop
  End If
End Sub

Private Sub txtNafn_LostFocus()
  Dim sL As String, sR As String
  Dim i As Integer, iL As Integer, iR As Integer, iT As Integer
  Dim iFjoldi As Integer
  Dim rsR As Recordset
  Dim result
  Dim strNafn As String

  txtNafn.Text = Trim(txtNafn.Text)
  strNafn = Replace(txtNafn.Text, "'", "''")
  For i = 1 To Len(strNafn)
    iT = Asc(Mid(strNafn, i, 1))
    If iT < 65 Or iT > 122 Or (iT > 90 And iT < 97) Then
      Mid(strNafn, i, 1) = "*"
    End If
  Next
  

' Giska á annan hund
  iFjoldi = 0
  If m_nr = 0 Then
    Set rsR = m_db.OpenRecordset("select * from hundalisti where nafn like '" & strNafn & "' and tegundaflokkur='" & pStillingar.Tegundaflokkur & "'")
    If Not rsR.EOF Then
      rsR.MoveLast
      iFjoldi = rsR.RecordCount
      rsR.MoveFirst
    End If
    If iFjoldi = 0 Then
      Set rsR = m_db.OpenRecordset("select * from hundalisti where nafn like '*" & strNafn & "' and tegundaflokkur='" & pStillingar.Tegundaflokkur & "'")
      If Not rsR.EOF Then
        rsR.MoveLast
        iFjoldi = rsR.RecordCount
        rsR.MoveFirst
      End If
    End If
    If iFjoldi = 0 Then
      Set rsR = m_db.OpenRecordset("select * from hundalisti where nafn like '*" & strNafn & "*' and tegundaflokkur='" & pStillingar.Tegundaflokkur & "'")
      If Not rsR.EOF Then
        rsR.MoveLast
        iFjoldi = rsR.RecordCount
        rsR.MoveFirst
      End If
    End If
    
    If iFjoldi <= 4 Then
      While Not rsR.EOF And result <> 6
        result = MsgBox("Er " & rsR!Nafn & "," & rsR!aettbokarnr & "," & rsR!fdags & " sami hundurinn?", vbYesNo, "Sami hundur") 'sL & ":" & sR
        If result = 6 Then
          m_nr = rsR!nr
          m_hladid = False
          Form_Load
          Form_activate
        End If
        rsR.MoveNext
      Wend
    End If
    rsR.Close
  End If
' Giska á ræktanda

'  If txtRaektandi.Tag <> 0 Then Exit Sub
'
'  iL = InStr(txtNafn.Text, " ")
'  iT = InStr(txtNafn.Text, "-")
'  If iL = 0 And iT > 0 Then iL = iT
'  If iT < iL And iT <> 0 Then iL = iT
'  iT = InStr(txtNafn.Text, " ")
'  While iT > 1
'    iR = iT
'    iT = InStr(iT + 1, txtNafn.Text, " ")
'  Wend
'  If iL = 0 And iR = 0 Then Exit Sub
'  sL = Left(txtNafn.Text, iL - 1)
'  sL = Replace(sL, "'", "''")
'  sR = Right(txtNafn.Text, Len(txtNafn.Text) - iR)
'  sR = Replace(sR, "'", "''")
'
'  Set rsR = m_db.OpenRecordset("select * from eigendur where raektunarnafn like '" & sL & "*' or raektunarnafn like '*" & sR & "'")
'  While Not rsR.EOF
'    result = MsgBox("Er " & rsR!Nafn & "," & rsR!raektunarnafn & " ræktandinn?", vbYesNo, "Ræktandi") 'sL & ":" & sR
'    If result = 6 Then
'      txtRaektandi.Text = rsR!Nafn
'      txtRaektandi.Tag = rsR!id
'    End If
'    rsR.MoveNext
'  Wend
'  rsR.Close
End Sub
