VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00CDEBEB&
   Caption         =   "Eigendur"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   LinkTopic       =   "Form10"
   ScaleHeight     =   6750
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoka 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "Loka"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDEBEB&
      Height          =   5535
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton cmdEyda 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFCDCD&
         Caption         =   "Eyða"
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton cmdBreyta 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFCDCD&
         Caption         =   "Breyta"
         Height          =   375
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton cmdNyr 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFCDCD&
         Caption         =   "Nýr"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   5040
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00CDEBEB&
         Caption         =   "Ræktendur"
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00CDEBEB&
         Caption         =   "Eigendur"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.ListBox List1 
         Height          =   4155
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   4335
      End
   End
   Begin VB.CommandButton cmdVistaHund 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "Nota"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDEBEB&
      Height          =   6495
      Left            =   4800
      TabIndex        =   12
      Top             =   120
      Width           =   4575
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   4080
         Width           =   4335
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   5520
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   4800
         Width           =   4335
      End
      Begin VB.CommandButton cmdHaettavid 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFCDCD&
         Caption         =   "Hætta við"
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6000
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdSkra 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFCDCD&
         Caption         =   "Vista"
         Height          =   375
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6000
         Width           =   1095
      End
      Begin VB.CheckBox chkRaektandi 
         BackColor       =   &H00CDEBEB&
         Caption         =   "Ræktandi"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtStadur 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   2640
         Width           =   4335
      End
      Begin VB.TextBox txtLand 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   3360
         Width           =   4335
      End
      Begin VB.TextBox txtHeimili 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1920
         Width           =   4335
      End
      Begin VB.TextBox txtRaektunarnafn 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txtNafn 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Sími:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Vefsíða:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Ræktunarnafn:"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Staður:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Land:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Heimili:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nafn:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_db As Database
Dim m_skrarNafn As String
Dim pStillingar As CStillingar
Dim bEigandi As Boolean
Dim m_EigandiID As Long
Dim m_Nafn As String

Public Sub Grunnur(fdb As Database)
  Set m_db = fdb
End Sub

Public Sub Eigandi()
  bEigandi = True
  Option1.Value = True
End Sub

Public Sub Raektandi()
  bEigandi = False
  Option2.Value = True
End Sub

Public Function Nafn() As String
  Nafn = m_Nafn
End Function

Public Property Get EigandiID() As Long
  EigandiID = m_EigandiID
End Property

Public Property Let EigandiID(ByVal fID As Long)
  m_EigandiID = fID
End Property

Private Sub cmdBreyta_Click()
  Dim rs As Recordset
  m_EigandiID = List1.ItemData(List1.ListIndex)
  Set rs = m_db.OpenRecordset("select * from eigendur where id=" & m_EigandiID)
  If Not rs.EOF Then
    txtNafn.Text = "" & rs!Nafn
    txtRaektunarnafn.Text = "" & rs!raektunarnafn
    txtHeimili.Text = "" & rs!heimili
    txtStadur.Text = "" & rs!stadur
    'txtKennitala.Text = "" & rs!kennitala
    chkRaektandi.Value = IIf(rs!Raektandi = "já", 1, 0)
    txtLand.Text = "" & rs!land
  End If
  rs.Close
  'Frame1.Visible = False
  Frame2.Visible = True
  'cmdVistaHund.Visible = False
End Sub

Private Sub cmdEyda_Click()
  Dim rs As Recordset
  Dim result
  m_EigandiID = List1.ItemData(List1.ListIndex)
  Set rs = m_db.OpenRecordset("select * from hundar where eigandiid=" & m_EigandiID & " or raektandiid=" & m_EigandiID)
  If Not rs.EOF Then
    MsgBox "Ekki hægt að eyða eiganda/ræktanda" & vbCrLf & "Eigandi/ræktandi er skráður á " & rs.RecordCount & " hunda.", , "Eyða eiganda/ræktanda"
    rs.Close
    Exit Sub
  End If
  rs.Close
  
  result = MsgBox("Viltu örugglega eyða færslunni: " & vbCrLf & List1.Text & "?", vbYesNo, "Eyða eiganda/ræktanda")
  If result = 7 Then
    Exit Sub
  End If
  
  Set rs = m_db.OpenRecordset("select * from eigendur where id=" & m_EigandiID)
  If Not rs.EOF Then
    rs.Delete
  End If
  rs.Close
  m_EigandiID = m_EigandiID - 1
  Birta
End Sub

Private Sub cmdHaettavid_Click()
'  Frame1.Visible = True
'  Frame2.Visible = False
'  cmdVistaHund.Visible = True
'  hreinsa
End Sub

Private Sub cmdNyr_Click()
  m_EigandiID = 0
  hreinsa
  'Frame1.Visible = False
  Frame2.Visible = True
  'cmdVistaHund.Visible = False
End Sub

Private Sub cmdSkra_Click()
  Vista m_EigandiID
  cmdHaettavid_Click
  Birta
End Sub

Private Sub cmdVistaHund_Click()
  If Frame1.Visible And List1.ListIndex >= 0 Then
    m_EigandiID = List1.ItemData(List1.ListIndex)
    m_Nafn = List1.Text
  'Else
  '  m_EigandiID = VistaNyjan
  '  m_Nafn = txtNafn.Text
  End If
  cmdLoka_Click
End Sub

Private Sub Vista(nr As Long)
  Dim max As Long
  Dim rs As Recordset
  
  If nr = 0 Then
    Set rs = m_db.OpenRecordset("select max(id) as maxid from eigendur")
    If Not rs.EOF Then
      m_EigandiID = Val("" & rs!maxid) + 1
    Else
      m_EigandiID = 1
    End If
    rs.Close
    
    Set rs = m_db.OpenRecordset("eigendur")
    rs.AddNew
    rs!id = m_EigandiID
  Else
    m_EigandiID = nr
    
    Set rs = m_db.OpenRecordset("select * from eigendur where id=" & m_EigandiID)
    rs.Edit
  End If
  
  rs!Nafn = txtNafn.Text
  rs!raektunarnafn = txtRaektunarnafn.Text
  rs!heimili = txtHeimili.Text
  rs!stadur = txtStadur.Text
  'rs!kennitala = txtKennitala.Text
  rs!Raektandi = IIf(chkRaektandi.Value = 1, "já", "-")
  rs!land = txtLand.Text
  rs.Update
  rs.Close
  m_Nafn = txtNafn.Text
End Sub

'Private Sub Command1_Click()
'  Frame1.Visible = Not Frame1.Visible
'  Frame2.Visible = Not Frame2.Visible
'  If bEigandi Then
'    Command1.Caption = IIf(Frame1.Visible, "Nýr eigandi", "Finna eiganda")
'  Else
'    Command1.Caption = IIf(Frame1.Visible, "Nýr ræktandi", "Finna ræktanda")
'  End If
'End Sub

Private Sub cmdLoka_Click()
  Me.Hide
End Sub

'Private Sub Command2_Click()
'  m_EigandiID = VistaNyjan
'  m_Nafn = txtNafn.Text
  'Command1_Click
'  Birta
'End Sub

Private Sub Form_activate()
  Birta
End Sub

Sub Birta()
  Dim rs As Recordset
  
  m_Nafn = ""
  
  Frame1.Visible = True
  'Frame2.Visible = False
  If bEigandi Then
    'Command1.Caption = IIf(Frame1.Visible, "Nýr eigandi", "Finna eiganda")
    Form10.Caption = "Eigendur"
  Else
    'Command1.Caption = IIf(Frame1.Visible, "Nýr ræktandi", "Finna ræktanda")
    Form10.Caption = "Ræktendur"
  End If
  
  List1.Clear
  List1.AddItem "(Óþekktur ræktandi)"
  List1.ItemData(List1.ListCount - 1) = 0
  If bEigandi Then
    Set rs = m_db.OpenRecordset("select * from eigendur order by raektunarnafn & nafn")
  Else
    Set rs = m_db.OpenRecordset("select * from eigendur where raektandi='já' order by raektunarnafn, nafn")
  End If
  While Not rs.EOF
    List1.AddItem IIf("" & rs!raektunarnafn <> "", rs!raektunarnafn & ", ", "") & rs!Nafn & " " & rs!heimili & " " & rs!stadur
    List1.ItemData(List1.ListCount - 1) = rs!id
    If EigandiID = rs!id Then List1.ListIndex = List1.ListCount - 1
    rs.MoveNext
  Wend
  rs.Close
  
End Sub

Private Sub List1_Click()
  If List1.ListIndex = 0 Then
    cmdBreyta.Enabled = False
    cmdEyda.Enabled = False
  Else
    cmdBreyta.Enabled = True
    cmdEyda.Enabled = True
  End If
End Sub

Private Sub List1_DblClick()
  cmdBreyta_Click
End Sub

Private Sub Option1_Click()
  bEigandi = True
  Option1.Value = True
  Birta
End Sub

Private Sub Option2_Click()
  bEigandi = False
  Option2.Value = True
  Birta
End Sub

Private Sub txtRaektunarnafn_Change()
  If txtRaektunarnafn <> "" Then
    chkRaektandi.Value = 1
    chkRaektandi.Enabled = False
  Else
    chkRaektandi.Enabled = True
  End If
End Sub

Sub hreinsa()
  txtNafn.Text = ""
  txtRaektunarnafn.Text = ""
  txtHeimili.Text = ""
  txtStadur.Text = ""
  'txtKennitala.Text = ""
  chkRaektandi.Value = 0
  txtLand.Text = ""
End Sub
