VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H00CDEBEB&
   Caption         =   "Laga tvískráningu"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4005
   LinkTopic       =   "Form5"
   ScaleHeight     =   3330
   ScaleWidth      =   4005
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLeit 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   3735
   End
   Begin VB.CommandButton cmdVistaHund 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "Laga"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoka 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "Loka"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtNafn 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3735
   End
   Begin VB.ComboBox cboF 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Leitarskilyrði:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Hundurinn:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Er sá sami og:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_db As Database
Private pStillingar As CStillingar

Dim m_nr As Long

Public Sub Stillingar(fStillingar As CStillingar)
  Set pStillingar = fStillingar
End Sub

Property Get nr() As Integer
  nr = m_nr
End Property

Public Property Let nr(ByVal vNewValue As Integer)
  m_nr = vNewValue
End Property

Sub Grunnur(fdb As Database)
  Set m_db = fdb
End Sub

Sub Nafn(fnafn As String)
  txtNafn.Text = fnafn
End Sub

Private Sub Form_activate()
  Dim rs As Recordset
  
  cboF.Clear
  Set rs = m_db.OpenRecordset("select * from hundalisti where nr<>" & m_nr & " and nafn like '*" & txtLeit.Text & "*' and latinn<>'fram' and tegundid=" & pStillingar.Tegund & " order by nafn, fdags")
  While Not rs.EOF
    cboF.AddItem IIf(rs!latinn = "fram", "* ", "") & rs!Nafn & " - " & rs!fdags 'IIf(rs!titill <> "", rs!titill & " ", "") &
    cboF.ItemData(cboF.ListCount - 1) = rs!nr
    cboF.ListIndex = 0
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
  Dim nr As Integer
  Dim i As Integer
  
  nr = m_nr
  m_nr = cboF.ItemData(cboF.ListIndex)
  
  m_db.Execute ("update hundar set fadirnr=" & m_nr & " where fadirnr=" & nr)
  m_db.Execute ("update hundar set modirnr=" & m_nr & " where modirnr=" & nr)
  m_db.Execute ("delete from hundar where nr=" & nr)
  
'  For i = 1 To Val(Text2.Text) + Val(Text3.Text)
'    Dim max As Long
'    Set rs = pHundur.faGrunn.OpenRecordset("select max(nr) as maxnr from hundar")
'    max = rs.Fields("maxnr")
'    rs.Close
'    Set rs = pHundur.faGrunn.OpenRecordset("hundar")
'    rs.AddNew
'    m_nr = max + 1
'    rs.Fields("nr") = m_nr
'
'    rs.Fields("Titill") = ""
'    rs.Fields("Fdags") = txtFdags.Text
'    rs.Fields("AettbokarNr") = ""
'    rs.Fields("Innkallsnafn") = ""
'    rs.Fields("Mynd1") = ""
'    rs.Fields("Mynd2") = ""
'
'    rs.Fields("latinn") = "-"
'    rs.Fields("geldur") = "-"
'
'    If i <= Val(Text2.Text) Then
'      rs.Fields("kyn") = "kk"
'      rs.Fields("Nafn") = "Hundur" & str(i)
'    Else
'      rs.Fields("kyn") = "kvk"
'      rs.Fields("Nafn") = "Tík" & str(i - Val(Text2.Text))
'    End If
'
'    rs.Fields("fadirnr") = cboF.ItemData(cboF.ListIndex)
'    rs.Fields("modirnr") = cboM.ItemData(cboM.ListIndex)
'
'    If (cboR.ListIndex <> -1) Then
'      rs.Fields("raektandiid") = cboR.ItemData(cboR.ListIndex)
'    End If
'    On Error Resume Next
'    rs.Fields("landid") = cboLand.ItemData(cboLand.ListIndex)
'    On Error GoTo 0
'    rs.Fields("tegundid") = cboTegund.ItemData(cboTegund.ListIndex)
'
'    rs.Fields("eiginhundar") = "-"
'    rs.Fields("klubbhundar") = IIf(chkKlubbhundur.Value = 1, "já", "-")
'    rs.Fields("minraektun") = IIf(chkMinRaektun.Value = 1, "já", "-")
'    rs.Fields("latinn") = IIf(chkFramraektun.Value = 1, "fram", "-")
'
'    rs.Update
'    rs.Close
'  Next
  MsgBox "Tvískráning lagfærð", vbOKOnly, "Tvískráning"
  cmdLoka_Click
End Sub


Private Sub txtLeit_Change()
 Form_activate
End Sub
