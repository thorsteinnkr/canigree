VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form9 
   BackColor       =   &H00CDEBEB&
   Caption         =   "Hundar"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   LinkTopic       =   "Form9"
   ScaleHeight     =   4440
   ScaleWidth      =   4710
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtVolume 
      Height          =   285
      Left            =   240
      TabIndex        =   16
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtLeyfi 
      Height          =   285
      Left            =   1200
      TabIndex        =   14
      Top             =   3600
      Width           =   3255
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Top             =   3240
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00AFCDCD&
      Caption         =   "..."
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6000
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   600
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00AFCDCD&
      Caption         =   ">>>"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cboTegund 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2880
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   6000
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   5520
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   5040
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00AFCDCD&
      Caption         =   "Áfram"
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Leyfisnúmer"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   3240
      Width           =   855
   End
   Begin VB.Shape Shape1 
      Height          =   1575
      Left            =   120
      Top             =   4920
      Width           =   4455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tegund"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Gagnaskrá"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mappa með myndum"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mappa fyrir ættbækur"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   5160
      Width           =   855
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_db As Database
Private pStillingar As CStillingar
Private m_hladid As Boolean

Public Sub Stillingar(fStillingar As CStillingar)
  Set pStillingar = fStillingar
End Sub

Private Sub Command1_Click()
  Dim rs As Recordset
  pStillingar.Mappa = Text2.Text
  pStillingar.Myndir = Text3.Text
  pStillingar.Grunnur = Text4.Text
  pStillingar.Tegund = cboTegund.ItemData(cboTegund.ListIndex)
  Set rs = m_db.OpenRecordset("select fciflokkur from tegundir where id=" & cboTegund.ItemData(cboTegund.ListIndex))
  If Not rs.EOF Then
    pStillingar.Tegundaflokkur = rs!fciflokkur
  End If
  rs.Close
  pStillingar.Tegundheiti = cboTegund.Text
  pStillingar.Email = txtEmail.Text
  pStillingar.Leyfi = txtLeyfi.Text
  'pStillingar.Cpuid = txtCpuid.Text
  pStillingar.ReCreate
  Me.Hide
End Sub

Private Sub Command2_Click()
  If Form9.Height = 6075 Then
    Form9.Height = 4845 '3795
    Command2.Caption = ">>>"
  Else
    Form9.Height = 7095
    Command2.Caption = "<<<"
  End If
End Sub

Private Sub Command3_Click()
  Dim result
  cdl.FileName = pStillingar.Grunnur
  cdl.Filter = "Gagnaskrár (*.mdb)|*.mdb"
  cdl.ShowOpen
  Text4.Text = cdl.FileName
End Sub

Private Sub Form_activate()
  If Not m_hladid Then
    Text2.Text = pStillingar.Mappa
    Text3.Text = pStillingar.Myndir
    Text4.Text = pStillingar.Grunnur
    txtEmail.Text = pStillingar.Email
    
    
    
    txtLeyfi.Text = pStillingar.Leyfi
    synaTegundir
  End If
  m_hladid = True
End Sub

Sub synaTegundir()
  On Error GoTo endir
  Dim rs As Recordset
  cboTegund.Clear
  Set m_db = OpenDatabase(pStillingar.Grunnur)
  Set rs = m_db.OpenRecordset("select * from tegundir order by heiti")
  While Not rs.EOF
    cboTegund.AddItem rs!heiti
    cboTegund.ItemData(cboTegund.ListCount - 1) = rs!id
    If rs!id = pStillingar.Tegund Then cboTegund.ListIndex = cboTegund.ListCount - 1
    rs.MoveNext
  Wend
  rs.Close
  Exit Sub
endir:
  MsgBox "Gagnaskrá er í ólagi, athugaðu skráninguna.", vbOKOnly, "Villa"
End Sub

Private Sub Form_Load()
  Text1.Text = "CaniGree - Ættbókarforrit fyrir hunda." & vbCrLf
  Text1.Text = Text1.Text & "Höfundarréttur © 2000-" & Year(Now()) & " Þorsteinn Kristinsson" & vbCrLf & vbCrLf
  Text1.Text = Text1.Text & "Afritun og dreifing á þessu tölvuforriti er óheimil með öllu nema með leyfi höfundar." & vbCrLf & vbCrLf
  Text1.Text = Text1.Text & "Nánari upplýsingar er að fá á heimasíðu forritsins http://www.canigree.com.  Tilkynningar um villur/bilanir eða tillögur að viðbótum skal senda á: tk@canigree.com."
  Form9.Height = 4845 '3795
  Command2.Caption = ">>>"
End Sub

Private Sub Text4_LostFocus()
  pStillingar.Grunnur = Text4.Text
  synaTegundir
End Sub
