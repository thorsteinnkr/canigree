VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00CDEBEB&
   Caption         =   "Ættbók"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10680
   LinkTopic       =   "Form4"
   ScaleHeight     =   7200
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBreytaMM 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "..."
      Height          =   255
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6000
      Width           =   255
   End
   Begin VB.CommandButton cmdBreytaFM 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "..."
      Height          =   255
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton cmdBreytaMF 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "..."
      Height          =   255
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton cmdBreytaFF 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "..."
      Height          =   255
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton cmdBreyta 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "..."
      Height          =   255
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4200
      Width           =   255
   End
   Begin VB.CommandButton cmdBreytaF 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "..."
      Height          =   255
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton cmdBreytaM 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "..."
      Height          =   255
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5760
      Width           =   255
   End
   Begin VB.CommandButton cmdMMM 
      BackColor       =   &H00AFCDCD&
      Caption         =   "Nýskrá hund"
      Height          =   375
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdFMM 
      BackColor       =   &H00AFCDCD&
      Caption         =   "Nýskrá hund"
      Height          =   375
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdMFM 
      BackColor       =   &H00AFCDCD&
      Caption         =   "Nýskrá hund"
      Height          =   375
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdFFM 
      BackColor       =   &H00AFCDCD&
      Caption         =   "Nýskrá hund"
      Height          =   375
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdMMF 
      BackColor       =   &H00AFCDCD&
      Caption         =   "Nýskrá hund"
      Height          =   375
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdFMF 
      BackColor       =   &H00AFCDCD&
      Caption         =   "Nýskrá hund"
      Height          =   375
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdMFF 
      BackColor       =   &H00AFCDCD&
      Caption         =   "Nýskrá hund"
      Height          =   375
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdFFF 
      BackColor       =   &H00AFCDCD&
      Caption         =   "Nýskrá hund"
      Height          =   375
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdMM 
      BackColor       =   &H00AFCDCD&
      Caption         =   "Nýskrá hund"
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdFM 
      BackColor       =   &H00AFCDCD&
      Caption         =   "Nýskrá hund"
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdMF 
      BackColor       =   &H00AFCDCD&
      Caption         =   "Nýskrá hund"
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdFF 
      BackColor       =   &H00AFCDCD&
      Caption         =   "Nýskrá hund"
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdM 
      BackColor       =   &H00AFCDCD&
      Caption         =   "Nýskrá hund"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdF 
      BackColor       =   &H00AFCDCD&
      Caption         =   "Nýskrá hund"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstAfkvaemi 
      Height          =   1620
      Left            =   240
      TabIndex        =   16
      Top             =   4800
      Width           =   2415
   End
   Begin VB.TextBox txtNafn 
      Height          =   2415
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtMMM 
      Height          =   615
      Left            =   8040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   5880
      Width           =   2415
   End
   Begin VB.TextBox txtMM 
      Height          =   1095
      Left            =   5520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   5160
      Width           =   2415
   End
   Begin VB.TextBox txtM 
      Height          =   1095
      Left            =   3000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox txtFM 
      Height          =   1095
      Left            =   5520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox txtMF 
      Height          =   1095
      Left            =   5520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton cmdLoka 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "Loka"
      Height          =   375
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox txtFMM 
      Height          =   615
      Left            =   8040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   5160
      Width           =   2415
   End
   Begin VB.TextBox txtMFM 
      Height          =   615
      Left            =   8040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox txtFFM 
      Height          =   615
      Left            =   8040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox txtMMF 
      Height          =   615
      Left            =   8040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox txtFMF 
      Height          =   615
      Left            =   8040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox txtMFF 
      Height          =   615
      Left            =   8040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtFFF 
      Height          =   615
      Left            =   8040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txtFF 
      Height          =   1095
      Left            =   5520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txtF 
      Height          =   1095
      Left            =   3000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      Height          =   5895
      Left            =   2880
      Top             =   720
      Width           =   7695
   End
   Begin VB.Label lblNafn 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   40
      Top             =   120
      Width           =   9855
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   240
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1335
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   1680
      Stretch         =   -1  'True
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Afkvæmi:"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lblAfkvaemi 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(0)"
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   4560
      Width           =   735
   End
   Begin VB.Image Image1M 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   975
   End
   Begin VB.Image Image2M 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Image Image1F 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   975
   End
   Begin VB.Image Image2F 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      Height          =   5895
      Left            =   120
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pHundur As CHundur
Private pStillingar As CStillingar

Sub hundur(fHundur As CHundur)
  Set pHundur = fHundur
End Sub

Public Sub Stillingar(fStillingar As CStillingar)
  Set pStillingar = fStillingar
End Sub

Sub SetjaF(fnr, fnrF)
  If fnrF = 0 Then Exit Sub
  Dim rs As Recordset
  Set rs = pHundur.faGrunn.OpenRecordset("select * from hundar where nr=" & fnr)
  If Not rs.EOF Then
    rs.Edit
    rs!fadirnr = fnrF
    rs.Update
  End If
  rs.Close
End Sub

Sub SetjaM(fnr, fnrM)
  If fnrM = 0 Then Exit Sub
  Dim rs As Recordset
  Set rs = pHundur.faGrunn.OpenRecordset("select * from hundar where nr=" & fnr)
  If Not rs.EOF Then
    rs.Edit
    rs!modirnr = fnrM
    rs.Update
  End If
  rs.Close
End Sub

Private Sub cmdBreytaFF_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.Stillingar pStillingar
  Form3.hundur IIf("" & pHundur.saekja("nrff") <> "", pHundur.saekja("nrff"), 0)
  Form3.Show vbModal
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub cmdBreytaFM_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.Stillingar pStillingar
  Form3.hundur IIf("" & pHundur.saekja("nrfm") <> "", pHundur.saekja("nrfm"), 0)
  Form3.Show vbModal
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub cmdBreytaMF_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.Stillingar pStillingar
  Form3.hundur IIf("" & pHundur.saekja("nrmf") <> "", pHundur.saekja("nrmf"), 0)
  Form3.Show vbModal
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub cmdBreytaMM_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.Stillingar pStillingar
  Form3.hundur IIf("" & pHundur.saekja("nrmm") <> "", pHundur.saekja("nrmm"), 0)
  Form3.Show vbModal
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub cmdF_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.hundur 0, "kk"
  Form3.Stillingar pStillingar
  Form3.Show vbModal
  SetjaF txtNafn.Tag, Form3.faHund
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub cmdFF_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.hundur 0, "kk"
  Form3.Stillingar pStillingar
  Form3.Show vbModal
  SetjaF txtF.Tag, Form3.faHund
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub cmdFFF_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.hundur 0, "kk"
  Form3.Stillingar pStillingar
  Form3.Show vbModal
  SetjaF txtFF.Tag, Form3.faHund
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub cmdFFM_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.hundur 0, "kk"
  Form3.Stillingar pStillingar
  Form3.Show vbModal
  SetjaF txtFM.Tag, Form3.faHund
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub cmdFM_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.hundur 0, "kk"
  Form3.Stillingar pStillingar
  Form3.Show vbModal
  SetjaF txtM.Tag, Form3.faHund
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub cmdFMF_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.hundur 0, "kk"
  Form3.Stillingar pStillingar
  Form3.Show vbModal
  SetjaF txtMF.Tag, Form3.faHund
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub cmdFMM_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.hundur 0, "kk"
  Form3.Stillingar pStillingar
  Form3.Show vbModal
  SetjaF txtMM.Tag, Form3.faHund
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub cmdM_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.hundur 0, "kvk"
  Form3.Stillingar pStillingar
  Form3.Show vbModal
  SetjaM txtNafn.Tag, Form3.faHund
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub cmdLoka_Click()
  Me.Hide
End Sub

Private Sub Command9_Click()

End Sub

Private Sub cmdBreytaM_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.Stillingar pStillingar
  Form3.hundur IIf("" & pHundur.saekja("nrm") <> "", pHundur.saekja("nrm"), 0)
  Form3.Show vbModal
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub cmdBreytaF_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.Stillingar pStillingar
  Form3.hundur IIf("" & pHundur.saekja("nrf") <> "", pHundur.saekja("nrf"), 0)
  Form3.Show vbModal
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub cmdBreyta_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.Stillingar pStillingar
  Form3.hundur IIf("" & pHundur.saekja("nr") <> "", pHundur.saekja("nr"), 0)
  Form3.Show vbModal
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub cmdMF_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.hundur 0, "kvk"
  Form3.Stillingar pStillingar
  Form3.Show vbModal
  SetjaM txtF.Tag, Form3.faHund
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub cmdMFF_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.hundur 0, "kvk"
  Form3.Stillingar pStillingar
  Form3.Show vbModal
  SetjaM txtFF.Tag, Form3.faHund
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub cmdMFM_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.hundur 0, "kvk"
  Form3.Stillingar pStillingar
  Form3.Show vbModal
  SetjaM txtFM.Tag, Form3.faHund
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub cmdMM_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.hundur 0, "kvk"
  Form3.Stillingar pStillingar
  Form3.Show vbModal
  SetjaM txtM.Tag, Form3.faHund
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub cmdMMF_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.hundur 0, "kvk"
  Form3.Stillingar pStillingar
  Form3.Show vbModal
  SetjaM txtMF.Tag, Form3.faHund
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub cmdMMM_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.hundur 0, "kvk"
  Form3.Stillingar pStillingar
  Form3.Show vbModal
  SetjaM txtMM.Tag, Form3.faHund
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub Form_activate()
  Birta
End Sub

Sub Birta()
  Dim mynd As String
  
  lblNafn.Caption = IIf(pHundur.saekja("titill") <> "", pHundur.saekja("titill") & " ", "") & pHundur.saekja("nafn")
  
  'txtNafn.Text = IIf(pHundur.saekja("titill") <> "", pHundur.saekja("titill") & " ", "") & pHundur.saekja("nafn") & " " & pHundur.saekja("aettbokarnr")
  'If pHundur.saekja("titill") <> "" Then txtNafn.ForeColor = &HFF Else txtNafn.ForeColor = &H0
  txtNafn.Tag = pHundur.saekja("nr")
  
  If "" & pHundur.saekja("titill") <> "" Then txtNafn.ForeColor = &HFF Else txtNafn.ForeColor = &H0
  txtNafn.Text = "Nafn: " & IIf(pHundur.saekja("titill") <> "", pHundur.saekja("titill") & " ", "") & pHundur.saekja("nafn") & vbCrLf
  txtNafn.Text = txtNafn.Text & IIf(pHundur.saekja("innkallsnafn") <> "", "Innkallsnafn: " & pHundur.saekja("innkallsnafn") & vbCrLf, "")
  txtNafn.Text = txtNafn.Text & "Fæðingardagur: " & pHundur.saekja("fdags") & vbCrLf
  txtNafn.Text = txtNafn.Text & "Ættbókarnúmer: " & pHundur.saekja("aettbokarnr") & vbCrLf
  txtNafn.Text = txtNafn.Text & "Örmerki: " & pHundur.saekja("ormerki") & vbCrLf
  txtNafn.Text = txtNafn.Text & "Litur: " & pHundur.saekja("litur") & vbCrLf
  txtNafn.Text = txtNafn.Text & "Kyn: " & IIf(pHundur.saekja("kyn") = "kk", "Karlhundur", IIf(pHundur.saekja("kyn") = "kvk", "Tík", "Ekki skráð")) & vbCrLf
  txtNafn.Text = txtNafn.Text & "Fæðingarstaður: " & pHundur.saekja("landheiti") & vbCrLf
  txtNafn.Text = txtNafn.Text & IIf("" & pHundur.saekja("raektandinafn") <> "", "Ræktandi: " & pHundur.saekja("raektandinafn") & vbCrLf, "")
  txtNafn.Text = txtNafn.Text & IIf("" & pHundur.saekja("eigandinafn") <> "", "Eigandi: " & pHundur.saekja("eigandinafn") & vbCrLf, "")
  If pHundur.saekja("geldur") = "já" Then
    txtNafn.Text = txtNafn.Text & "Geldur: já" & vbCrLf
  End If
  If pHundur.saekja("latinn") = "já" Then
    txtNafn.Text = txtNafn.Text & IIf(pHundur.saekja("kyn") = "kvk", "Látin: ", "Látinn: ") & "já" & vbCrLf
  End If
  
  
  
  txtF.Text = uppl("f")
  If pHundur.saekja("tf") <> "" Then txtF.ForeColor = &HFF Else txtF.ForeColor = &H0
  txtF.Tag = pHundur.saekja("nrf")
  cmdF.Visible = IIf(txtF.Tag = "", True, False)
  txtF.Visible = IIf(txtF.Tag <> "", True, False)
  cmdBreytaF.Visible = IIf(txtF.Tag <> "", True, False)
  
  txtM = uppl("m")
  If pHundur.saekja("tm") <> "" Then txtM.ForeColor = &HFF Else txtM.ForeColor = &H0
  txtM.Tag = pHundur.saekja("nrM")
  cmdM.Visible = IIf(txtM.Tag = "", True, False)
  txtM.Visible = IIf(txtM.Tag = "", False, True)
  cmdBreytaM.Visible = IIf(txtM.Tag <> "", True, False)
  
  txtFF = uppl("ff")
  If pHundur.saekja("tff") <> "" Then txtFF.ForeColor = &HFF Else txtFF.ForeColor = &H0
  txtFF.Tag = pHundur.saekja("nrff")
  cmdFF.Visible = IIf(txtFF.Tag = "" And txtF.Tag <> "", True, False)
  txtFF.Visible = IIf(txtFF.Tag = "" And txtF.Tag <> "", False, True)
  cmdBreytaFF.Visible = IIf(txtFF.Tag <> "", True, False)
  
  txtMF = uppl("mf")
  If pHundur.saekja("tmf") <> "" Then txtMF.ForeColor = &HFF Else txtMF.ForeColor = &H0
  txtMF.Tag = pHundur.saekja("nrmf")
  cmdMF.Visible = IIf(txtMF.Tag = "" And txtF.Tag <> "", True, False)
  txtMF.Visible = IIf(txtMF.Tag = "" And txtF.Tag <> "", False, True)
  cmdBreytaMF.Visible = IIf(txtMF.Tag <> "", True, False)

  txtFM = uppl("fm")
  If pHundur.saekja("tfm") <> "" Then txtFM.ForeColor = &HFF Else txtFM.ForeColor = &H0
  txtFM.Tag = pHundur.saekja("nrfm")
  cmdFM.Visible = IIf(txtFM.Tag = "" And txtM.Tag <> "", True, False)
  txtFM.Visible = IIf(txtFM.Tag = "" And txtM.Tag <> "", False, True)
  cmdBreytaFM.Visible = IIf(txtFM.Tag <> "", True, False)
  
  txtMM = uppl("mm")
  If pHundur.saekja("tmm") <> "" Then txtMM.ForeColor = &HFF Else txtMM.ForeColor = &H0
  txtMM.Tag = pHundur.saekja("nrmm")
  cmdMM.Visible = IIf(txtMM.Tag = "" And txtM.Tag <> "", True, False)
  txtMM.Visible = IIf(txtMM.Tag = "" And txtM.Tag <> "", False, True)
  cmdBreytaMM.Visible = IIf(txtMM.Tag <> "", True, False)
  
  txtFFF = uppl("fff")
  If pHundur.saekja("tfff") <> "" Then txtFFF.ForeColor = &HFF Else txtFFF.ForeColor = &H0
  txtFFF.Tag = pHundur.saekja("nrfff")
  If txtFFF.Tag = "" And txtFF.Tag <> "" Then cmdFFF.Visible = True: txtFFF.Visible = False
  cmdFFF.Visible = IIf(txtFFF.Tag = "" And txtFF.Tag <> "", True, False)
  txtFFF.Visible = IIf(txtFFF.Tag = "" And txtFF.Tag <> "", False, True)
  
  txtMFF = uppl("mff")
  If pHundur.saekja("tmff") <> "" Then txtMFF.ForeColor = &HFF Else txtMFF.ForeColor = &H0
  txtMFF.Tag = pHundur.saekja("nrmff")
  cmdMFF.Visible = IIf(txtMFF.Tag = "" And txtFF.Tag <> "", True, False)
  txtMFF.Visible = IIf(txtMFF.Tag = "" And txtFF.Tag <> "", False, True)
  
  txtFMF = uppl("fmf")
  If pHundur.saekja("tfmf") <> "" Then txtFMF.ForeColor = &HFF Else txtFMF.ForeColor = &H0
  txtFMF.Tag = pHundur.saekja("nrfmf")
  cmdFMF.Visible = IIf(txtFMF.Tag = "" And txtMF.Tag <> "", True, False)
  txtFMF.Visible = IIf(txtFMF.Tag = "" And txtMF.Tag <> "", False, True)
  
  txtMMF = uppl("mmf")
  If pHundur.saekja("tmmf") <> "" Then txtMMF.ForeColor = &HFF Else txtMMF.ForeColor = &H0
  txtMMF.Tag = pHundur.saekja("nrmmf")
  cmdMMF.Visible = IIf(txtMMF.Tag = "" And txtMF.Tag <> "", True, False)
  txtMMF.Visible = IIf(txtMMF.Tag = "" And txtMF.Tag <> "", False, True)
  
  txtFFM = uppl("ffm")
  If pHundur.saekja("tffm") <> "" Then txtFFM.ForeColor = &HFF Else txtFFM.ForeColor = &H0
  txtFFM.Tag = pHundur.saekja("nrffm")
  cmdFFM.Visible = IIf(txtFFM.Tag = "" And txtFM.Tag <> "", True, False)
  txtFFM.Visible = IIf(txtFFM.Tag = "" And txtFM.Tag <> "", False, True)
  
  txtMFM = uppl("mfm")
  If pHundur.saekja("tmfm") <> "" Then txtMFM.ForeColor = &HFF Else txtMFM.ForeColor = &H0
  txtMFM.Tag = pHundur.saekja("nrmfm")
  cmdMFM.Visible = IIf(txtMFM.Tag = "" And txtFM.Tag <> "", True, False)
  txtMFM.Visible = IIf(txtMFM.Tag = "" And txtFM.Tag <> "", False, True)
  
  txtFMM = uppl("fmm")
  If pHundur.saekja("tfmm") <> "" Then txtFMM.ForeColor = &HFF Else txtFMM.ForeColor = &H0
  txtFMM.Tag = pHundur.saekja("nrfmm")
  cmdFMM.Visible = IIf(txtFMM.Tag = "" And txtMM.Tag <> "", True, False)
  txtFMM.Visible = IIf(txtFMM.Tag = "" And txtMM.Tag <> "", False, True)
  
  txtMMM = uppl("mmm")
  If pHundur.saekja("tmmm") <> "" Then txtMMM.ForeColor = &HFF Else txtMMM.ForeColor = &H0
  txtMMM.Tag = pHundur.saekja("nrmmm")
  cmdMMM.Visible = IIf(txtMMM.Tag = "" And txtMM.Tag <> "", True, False)
  txtMMM.Visible = IIf(txtMMM.Tag = "" And txtMM.Tag <> "", False, True)
  
  On Error Resume Next
  
  Image1 = LoadPicture
  mynd = pHundur.saekja("mynd1")
  If mynd <> "" Then
    Image1 = LoadPicture(pStillingar.Mappa & pStillingar.Myndir & "\" & mynd)
  End If
  
  Image2 = LoadPicture
  mynd = pHundur.saekja("mynd2")
  If mynd <> "" Then
    Image2 = LoadPicture(pStillingar.Mappa & pStillingar.Myndir & "\" & mynd)
  End If
  
  
  Image1F = LoadPicture
  mynd = pHundur.saekja("my1f")
  If mynd <> "" Then
    Image1F = LoadPicture(pStillingar.Mappa & pStillingar.Myndir & "\" & mynd)
  End If
  
  Image2F = LoadPicture
  mynd = pHundur.saekja("my2f")
  If mynd <> "" Then
    Image2F = LoadPicture(pStillingar.Mappa & pStillingar.Myndir & "\" & mynd)
  End If

  Image1M = LoadPicture
  mynd = pHundur.saekja("my1m")
  If mynd <> "" Then
    Image1M = LoadPicture(pStillingar.Mappa & pStillingar.Myndir & "\" & mynd)
  End If

  Image2M = LoadPicture
  mynd = pHundur.saekja("my2m")
  If mynd <> "" Then
    Image2M = LoadPicture(pStillingar.Mappa & pStillingar.Myndir & "\" & mynd)
  End If
  On Error GoTo 0

  BirtaAfkvaemi
  BirtaSystkini
End Sub

Private Sub BirtaAfkvaemi()
  lstAfkvaemi.Clear
  
  If "" & pHundur.saekja("nr") = "" Then Exit Sub
  
  pHundur.setjaLista ("nrf=" & pHundur.saekja("nr") & " or nrm=" & pHundur.saekja("nr") & " order by fdags")
  Do
    lstAfkvaemi.AddItem IIf(pHundur.saekjaLista("titill") <> "", pHundur.saekjaLista("titill") & " ", "") & pHundur.saekjaLista("nafn") & " - " & pHundur.saekjaLista("FDags")
    lstAfkvaemi.ItemData(lstAfkvaemi.ListCount - 1) = "0" & pHundur.saekjaLista("nr")
  Loop Until pHundur.naestaLista
  lblAfkvaemi = "(" & pHundur.fjoldiLista & ")"
  pHundur.lokaLista
End Sub

Private Sub BirtaSystkini()
  Dim s As String
  'lstSystkini.Clear
  
  If "" & pHundur.saekja("nrf") <> "" And "" & pHundur.saekja("nrm") <> "" Then
    pHundur.setjaLista ("(nrf=" & pHundur.saekja("nrf") & " or nrm=" & pHundur.saekja("nrm") & ") and nr<>" & pHundur.saekja("nr") & " order by fdags")
  ElseIf "" & pHundur.saekja("nrf") = "" And "" & pHundur.saekja("nrm") <> "" Then
    pHundur.setjaLista ("(nrm=" & pHundur.saekja("nrm") & ") and nr<>" & pHundur.saekja("nr") & " order by fdags")
  ElseIf "" & pHundur.saekja("nrf") <> "" And "" & pHundur.saekja("nrm") = "" Then
    pHundur.setjaLista ("(nrf=" & pHundur.saekja("nrf") & ") and nr<>" & pHundur.saekja("nr") & " order by fdags")
  Else
    Exit Sub
  End If
  Do
    s = ""
    If pHundur.saekjaLista("nrf") = pHundur.saekja("nrf") And pHundur.saekjaLista("nrm") = pHundur.saekja("nrm") Then
      s = "Alsystkin"
    ElseIf pHundur.saekjaLista("nrf") = pHundur.saekja("nrf") Then
      s = "Samfeðra"
    ElseIf pHundur.saekjaLista("nrm") = pHundur.saekja("nrm") Then
      s = "Sammæðra"
    End If

'    lstSystkini.AddItem IIf(pHundur.saekjaLista("titill") <> "", pHundur.saekjaLista("titill") & " ", "") & pHundur.saekjaLista("nafn") & " - " & pHundur.saekjaLista("FDags") & " " & s
'    lstSystkini.ItemData(lstSystkini.ListCount - 1) = "0" & pHundur.saekjaLista("nr")
  Loop Until pHundur.naestaLista
'  lblSystkini = "(" & pHundur.fjoldiLista & ")"
  pHundur.lokaLista
End Sub

Function uppl(ai As String)
  uppl = IIf(pHundur.saekja("t" & ai) <> "", pHundur.saekja("t" & ai) & " " & pHundur.saekja("n" & ai), pHundur.saekja("n" & ai)) & vbCrLf & pHundur.saekja("fd" & ai)
End Function

Private Sub Image1F_Click()
  Load Form2
  Form2.Image1 = Image1F
  Form2.Width = Form2.Image1.Width + 400
  Form2.Height = Form2.Image1.Height + 1150
  Form2.Show vbModal
End Sub

Private Sub Image2F_Click()
  Load Form2
  Form2.Image1 = Image2F
  Form2.Width = Form2.Image1.Width + 400
  Form2.Height = Form2.Image1.Height + 1150
  Form2.Show vbModal
End Sub

Private Sub Image1M_Click()
  Load Form2
  Form2.Image1 = Image1M
  Form2.Width = Form2.Image1.Width + 400
  Form2.Height = Form2.Image1.Height + 1150
  Form2.Show vbModal
End Sub

Private Sub Image2M_Click()
  Load Form2
  Form2.Image1 = Image2M
  Form2.Width = Form2.Image1.Width + 400
  Form2.Height = Form2.Image1.Height + 1150
  Form2.Show vbModal
End Sub

Private Sub Image1_Click()
  Load Form2
  Form2.Image1 = Image1
  Form2.Width = Form2.Image1.Width + 400
  Form2.Height = Form2.Image1.Height + 1150
  Form2.Show vbModal
End Sub

Private Sub Image2_Click()
  Load Form2
  Form2.Image1 = Image2
  Form2.Width = Form2.Image1.Width + 400
  Form2.Height = Form2.Image1.Height + 1150
  Form2.Show vbModal
End Sub

Private Sub lstAfkvaemi_DblClick()
  pHundur.setjaHund lstAfkvaemi.ItemData(lstAfkvaemi.ListIndex)
  Birta
End Sub

'Private Sub lstSystkini_DblClick()
  'pHundur.setjaHund lstSystkini.ItemData(lstSystkini.ListIndex)
'  Birta
'End Sub

Private Sub txtF_DblClick()
  If txtF.Tag = "" Then Exit Sub
  pHundur.setjaHund txtF.Tag
  Birta
End Sub

Private Sub txtM_DblClick()
  If txtM.Tag = "" Then Exit Sub
  pHundur.setjaHund txtM.Tag
  Birta
End Sub

Private Sub txtMF_DblClick()
  If txtMF.Tag = "" Then Exit Sub
  pHundur.setjaHund txtMF.Tag
  Birta
End Sub

Private Sub txtFF_DblClick()
  If txtFF.Tag = "" Then Exit Sub
  pHundur.setjaHund txtFF.Tag
  Birta
End Sub

Private Sub txtFM_DblClick()
  If txtFM.Tag = "" Then Exit Sub
  pHundur.setjaHund txtFM.Tag
  Birta
End Sub

Private Sub txtMM_DblClick()
  If txtMM.Tag = "" Then Exit Sub
  pHundur.setjaHund txtMM.Tag
  Birta
End Sub

Private Sub txtFFF_DblClick()
  If txtFFF.Tag = "" Then Exit Sub
  pHundur.setjaHund txtFFF.Tag
  Birta
End Sub

Private Sub txtMFF_DblClick()
  If txtMFF.Tag = "" Then Exit Sub
  pHundur.setjaHund txtMFF.Tag
  Birta
End Sub

Private Sub txtFMF_DblClick()
  If txtFMF.Tag = "" Then Exit Sub
  pHundur.setjaHund txtFMF.Tag
  Birta
End Sub

Private Sub txtMMF_DblClick()
  If txtMMF.Tag = "" Then Exit Sub
  pHundur.setjaHund txtMMF.Tag
  Birta
End Sub

Private Sub txtFFM_DblClick()
  If txtFFM.Tag = "" Then Exit Sub
  pHundur.setjaHund txtFFM.Tag
  Birta
End Sub

Private Sub txtMFM_DblClick()
  If txtMFM.Tag = "" Then Exit Sub
  pHundur.setjaHund txtMFM.Tag
  Birta
End Sub

Private Sub txtFMM_DblClick()
  If txtFMM.Tag = "" Then Exit Sub
  pHundur.setjaHund txtFMM.Tag
  Birta
End Sub

Private Sub txtMMM_DblClick()
  If txtMMM.Tag = "" Then Exit Sub
  pHundur.setjaHund txtMMM.Tag
  Birta
End Sub

