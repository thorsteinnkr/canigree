VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00CDEBEB&
   Caption         =   "Skrá got"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3255
   LinkTopic       =   "Form5"
   ScaleHeight     =   5640
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSynaFram 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Sýna framræktun"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtRaektandi 
      Height          =   285
      Left            =   120
      TabIndex        =   22
      Top             =   3600
      Width           =   3015
   End
   Begin VB.CheckBox chkFramraektun 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Framræktun"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CheckBox chkSynaAlla 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00CDEBEB&
      Caption         =   "Sýna alla hunda"
      Height          =   255
      Left            =   1680
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdPorunarbok 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "Vista"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkKlubbhundur 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Deildarhundur"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CheckBox chkMinRaektun 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Mín ræktun"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4920
      Width           =   1815
   End
   Begin VB.ComboBox cboTegund 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   4200
      Width           =   2295
   End
   Begin VB.ComboBox cboLand 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton cmdVistaHund 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "Vista"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoka 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "Loka"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Text            =   "0"
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Text            =   "0"
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox txtFdags 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Text            =   "1.1.2003"
      Top             =   2160
      Width           =   1815
   End
   Begin VB.ComboBox cboM 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   3015
   End
   Begin VB.ComboBox cboF 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Ræktandi:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Tegund:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Fæðingarland:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tíkur:"
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Karlhundar:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Fæðingardagur:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Móðir:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Faðir:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private m_db As Database
Private pHundur As CHundur
Private pStillingar As CStillingar
Dim pListar As CListar

Dim m_nr As Long

'Sub Grunnur(fdb As Database)
'  Set m_db = fdb
'End Sub

Sub Listar(fListar As CListar)
  Set pListar = fListar
End Sub

Sub hundur(fHundur As CHundur)
  Set pHundur = fHundur
End Sub

Public Sub Stillingar(fStillingar As CStillingar)
  Set pStillingar = fStillingar
End Sub

Public Sub porunarbok()
  Form5.Caption = "Pörunarbók"
  cmdPorunarbok.Visible = True
  Label11.Visible = False
  Label3.Visible = False
  cboLand.Visible = False
  chkSynaAlla.Visible = True
  'chkSynaFram.Visible = True
  Form5.Height = 2445
End Sub

Private Sub chkSynaAlla_Click()
  Form_activate
End Sub

Private Sub cmdPorunarbok_Click()
  If cboF.ListIndex >= 0 And cboM.ListIndex >= 0 Then
    pListar.porunarbok cboF.ItemData(cboF.ListIndex), cboM.ItemData(cboM.ListIndex)
  End If
  cmdLoka_Click
End Sub

Private Sub Form_activate()
  Dim rs As Recordset
  
  cboF.Clear
  Set rs = pHundur.faGrunn.OpenRecordset("select * from hundalisti where kyn='kk' " & IIf(chkSynaAlla.Value = 0, " and (klubbhundar='já' or latinn='fram') ", "") & " and latinn<>'já' and geldur<>'já' and tegundaflokkur='" & pStillingar.Tegundaflokkur & "' order by nafn, fdags")
  While Not rs.EOF
    cboF.AddItem IIf(rs!latinn = "fram", "* ", "") & rs!Nafn & " - " & rs!fdags 'IIf(rs!titill <> "", rs!titill & " ", "") &
    cboF.ItemData(cboF.ListCount - 1) = rs!nr
    If pHundur.saekja("nr") = rs!nr Then cboF.ListIndex = cboF.ListCount - 1
    rs.MoveNext
  Wend
  rs.Close
  
  cboM.Clear
  Set rs = pHundur.faGrunn.OpenRecordset("select * from hundalisti where kyn='kvk' " & IIf(chkSynaAlla.Value = 0, " and (klubbhundar='já' or latinn='fram') ", "") & " and latinn<>'já' and geldur<>'já' and tegundaflokkur='" & pStillingar.Tegundaflokkur & "' order by nafn, fdags")
  While Not rs.EOF
    cboM.AddItem IIf(rs!latinn = "fram", "* ", "") & rs!Nafn & " - " & rs!fdags 'IIf(rs!titill <> "", rs!titill & " ", "") &
    cboM.ItemData(cboM.ListCount - 1) = rs!nr
    If pHundur.saekja("nr") = rs!nr Then cboM.ListIndex = cboM.ListCount - 1
    rs.MoveNext
  Wend
  rs.Close
  
'  cboR.Clear
'  Set rs = pHundur.faGrunn.OpenRecordset("select * from eigendur")
'  While Not rs.EOF
'    cboR.AddItem rs!Nafn & ", " & rs!raektunarnafn & " " & rs!heimili & " " & rs!stadur
'    cboR.ItemData(cboR.ListCount - 1) = rs!id
'    rs.MoveNext
'  Wend
'  rs.Close
  
  cboLand.Clear
  Set rs = pHundur.faGrunn.OpenRecordset("select * from land order by heiti")
  While Not rs.EOF
    cboLand.AddItem rs!heiti
    cboLand.ItemData(cboLand.ListCount - 1) = rs!id
    rs.MoveNext
  Wend
  rs.Close
  
  cboTegund.Clear
  Set rs = pHundur.faGrunn.OpenRecordset("select * from tegundir where fciflokkur='" & pStillingar.Tegundaflokkur & "'")
  While Not rs.EOF
    cboTegund.AddItem rs!heiti
    cboTegund.ItemData(cboTegund.ListCount - 1) = rs!id
    If pStillingar.Tegund = rs!id Then cboTegund.ListIndex = cboTegund.ListCount - 1
    rs.MoveNext
  Wend
  rs.Close
End Sub

Private Sub cmdLoka_Click()
  Me.Hide
End Sub

Private Sub cmdVistaHund_Click()
  MousePointer = 11
  Vista
  MousePointer = 0
End Sub

Private Sub Vista()
  Dim rs As Recordset
  Dim i As Integer
  
  For i = 1 To Val(Text2.Text) + Val(Text3.Text)
    Dim max As Long
    Set rs = pHundur.faGrunn.OpenRecordset("select max(nr) as maxnr from hundar")
    max = rs.Fields("maxnr")
    rs.Close
    Set rs = pHundur.faGrunn.OpenRecordset("hundar")
    rs.AddNew
    m_nr = max + 1
    rs.Fields("nr") = m_nr
    
    rs.Fields("Titill") = ""
    rs.Fields("Fdags") = txtFdags.Text
    rs.Fields("AettbokarNr") = ""
    rs.Fields("Innkallsnafn") = ""
    rs.Fields("Mynd1") = ""
    rs.Fields("Mynd2") = ""
  
    rs.Fields("latinn") = "-"
    rs.Fields("geldur") = "-"
  
    If i <= Val(Text2.Text) Then
      rs.Fields("kyn") = "kk"
      rs.Fields("Nafn") = "Hundur" & str(i)
    Else
      rs.Fields("kyn") = "kvk"
      rs.Fields("Nafn") = "Tík" & str(i - Val(Text2.Text))
    End If
    
    rs.Fields("fadirnr") = cboF.ItemData(cboF.ListIndex)
    rs.Fields("modirnr") = cboM.ItemData(cboM.ListIndex)
    
'    If (cboR.ListIndex <> -1) Then
'      rs.Fields("raektandiid") = cboR.ItemData(cboR.ListIndex)
'    End If
    rs.Fields("raektandi") = txtRaektandi.Text
    On Error Resume Next
    rs.Fields("landid") = cboLand.ItemData(cboLand.ListIndex)
    On Error GoTo 0
    rs.Fields("tegundid") = cboTegund.ItemData(cboTegund.ListIndex)
  
    rs.Fields("eiginhundar") = "-"
    rs.Fields("klubbhundar") = IIf(chkKlubbhundur.Value = 1, "já", "-")
    rs.Fields("minraektun") = IIf(chkMinRaektun.Value = 1, "já", "-")
    rs.Fields("latinn") = IIf(chkFramraektun.Value = 1, "fram", "-")
  
    rs.Update
    rs.Close
  Next
  MsgBox "Got skráð", vbOKOnly, "Gotskráning"
  cmdLoka_Click
End Sub

Private Sub Form_Load()
  txtFdags = Format(Now, "d.m.yyyy")
End Sub
