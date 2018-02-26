VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00CDEBEB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CaniGree"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11940
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11940
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdl 
      Left            =   2400
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtAthugasemd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   31
      Top             =   7080
      Width           =   8055
   End
   Begin VB.OptionButton optLeit4 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Framræktun"
      Height          =   255
      Left            =   240
      MaskColor       =   &H8000000F&
      TabIndex        =   30
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton Command11 
      Height          =   855
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Ættartré"
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      Height          =   855
      Left            =   7800
      Picture         =   "Form1.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Línuræktun"
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Height          =   855
      Left            =   9480
      Picture         =   "Form1.frx":29F4
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Eigendur ..."
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Height          =   855
      Left            =   6240
      Picture         =   "Form1.frx":49E6
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Ættartré"
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Height          =   855
      Left            =   5400
      Picture         =   "Form1.frx":69D8
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Skrifa ættbækur ..."
      Top             =   0
      Width           =   855
   End
   Begin VB.ListBox lstSaga 
      Height          =   1425
      Left            =   240
      TabIndex        =   24
      Top             =   6720
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Height          =   855
      Left            =   4440
      Picture         =   "Form1.frx":89CA
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Pörunarlisti"
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Height          =   855
      Left            =   3600
      Picture         =   "Form1.frx":A9BC
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Pörunarbók..."
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Height          =   855
      Left            =   2760
      Picture         =   "Form1.frx":C9AE
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Sjá ættbók"
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Height          =   855
      Left            =   1800
      Picture         =   "Form1.frx":E9A0
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Skrá got ..."
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   960
      Picture         =   "Form1.frx":10992
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Breyta skráningu ..."
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Height          =   855
      Left            =   120
      Picture         =   "Form1.frx":12984
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Skrá hund ..."
      Top             =   0
      Width           =   855
   End
   Begin VB.ListBox lstSystkini 
      Height          =   1620
      Left            =   8520
      TabIndex        =   16
      Top             =   5040
      Width           =   3255
   End
   Begin VB.ListBox lstAfkvaemi 
      Height          =   1620
      Left            =   6120
      TabIndex        =   15
      Top             =   5040
      Width           =   2295
   End
   Begin VB.TextBox txtLeita 
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton cmdLeita 
      BackColor       =   &H00AFCDCD&
      Caption         =   "Leita"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   855
   End
   Begin VB.OptionButton optLeit1 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Allir hundar"
      Height          =   255
      Left            =   240
      MaskColor       =   &H8000000F&
      TabIndex        =   1
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.OptionButton optLeit2 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Deildarhundar"
      Height          =   255
      Left            =   240
      MaskColor       =   &H8000000F&
      TabIndex        =   2
      Top             =   1800
      UseMaskColor    =   -1  'True
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.OptionButton optLeit3 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Eigin hundar"
      Height          =   255
      Left            =   240
      MaskColor       =   &H8000000F&
      TabIndex        =   3
      Top             =   2040
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.ListBox lstLeit 
      Height          =   3375
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox txtStori 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   6120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label lblLeyfi 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   8400
      Width           =   11415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Móðir:"
      Height          =   255
      Left            =   9480
      TabIndex        =   34
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Faðir:"
      Height          =   255
      Left            =   9480
      TabIndex        =   33
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Upplýsingar um hundinn:"
      Height          =   255
      Left            =   3720
      TabIndex        =   32
      Top             =   6720
      Width           =   3015
   End
   Begin VB.Label lblNafn 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   29
      Top             =   1080
      Width           =   5775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Síðustu uppflettingar:"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      Height          =   1935
      Left            =   120
      Top             =   6360
      Width           =   3375
   End
   Begin VB.Label lblModir 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9480
      TabIndex        =   22
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label lblFadir 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9480
      TabIndex        =   21
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label lblSystkini 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(0)"
      Height          =   255
      Left            =   11040
      TabIndex        =   20
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label lblAfkvaemi 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(0)"
      Height          =   255
      Left            =   7680
      TabIndex        =   19
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Systkini:"
      Height          =   255
      Left            =   8520
      TabIndex        =   18
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Afkvæmi:"
      Height          =   255
      Left            =   6120
      TabIndex        =   17
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Shape Shape6 
      Height          =   5295
      Left            =   120
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label lblLeit 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(0)"
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   2280
      Width           =   495
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Shape Shape4 
      Height          =   7335
      Left            =   3600
      Top             =   960
      Width           =   8295
   End
   Begin VB.Menu mnuVista 
      Caption         =   "&Vista"
      Begin VB.Menu mnuSkraningVistaCanigree 
         Caption         =   "Vista Canigree skrá ..."
      End
      Begin VB.Menu mnuLesaInnCanigreeSkra 
         Caption         =   "Lesa inn Canigree skrá ..."
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSkraningVistaPED 
         Caption         =   "Vista .PED skrá..."
      End
      Begin VB.Menu mnuHaetta 
         Caption         =   "&Hætta"
      End
   End
   Begin VB.Menu mnuSkraning 
      Caption         =   "&Skráning"
      Begin VB.Menu mnuSkraningNyrHundur 
         Caption         =   "&Nýr hundur..."
      End
      Begin VB.Menu mnuSkraningBreytaSkraningu 
         Caption         =   "&Breyta skráningu..."
      End
      Begin VB.Menu mnuSkraningSkraGot 
         Caption         =   "&Skrá got..."
      End
      Begin VB.Menu mnuSkraningEigendur 
         Caption         =   "&Eigendur..."
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuSkoda 
      Caption         =   "S&koða"
      Begin VB.Menu mnuSkodaAettbok 
         Caption         =   "Æ&ttbók"
      End
      Begin VB.Menu mnuSkodaPorun 
         Caption         =   "&Pörun..."
      End
      Begin VB.Menu mnuSkodaPorunarlisti 
         Caption         =   "Pö&runarlisti"
      End
      Begin VB.Menu mnuSkodaICgildi 
         Caption         =   "IC-gildi"
      End
      Begin VB.Menu mnuSkodaICforfedur 
         Caption         =   "IC-forfeður"
      End
   End
   Begin VB.Menu mnuListar 
      Caption         =   "&Listar"
      Begin VB.Menu mnuListarAettbaekur 
         Caption         =   "Æ&ttbækur..."
      End
      Begin VB.Menu mnuListarAettartre 
         Caption         =   "Ættar&tré"
      End
   End
   Begin VB.Menu mnuVefur 
      Caption         =   "&Vefur"
      Begin VB.Menu mnuVefurSkodaHund 
         Caption         =   "&Skoða hund"
      End
      Begin VB.Menu mnuVefurHundalisti 
         Caption         =   "&Hundalisti"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pStillingar As CStillingar
Dim pHundur As CHundur
Dim pListar As CListar
Const m_strUtg As String = "Útg. 1.3.6"
Private m_hladid As Boolean

Private Sub Command1_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.Stillingar pStillingar
  Form3.hundur IIf("" & pHundur.saekja("nr") <> "", pHundur.saekja("nr"), 0)
  Form3.Show vbModal
  pHundur.setjaHund
  Birta
  Unload Form3
End Sub

Private Sub Command10_Click()
  Load Form12
  Form12.Listar pListar
  Form12.hundur pHundur
  Form12.Show vbModal
  Unload Form12
End Sub

Private Sub Command2_Click()
  Load Form6
  Form6.hundur pHundur
  Form6.Listar pListar
  Form6.Show vbModal
  Unload Form6
End Sub

Private Sub Command3_Click()
  Load Form11
  Form11.hundur pHundur
  Form11.Listar pListar
  Form11.Show vbModal
  Unload Form11
'  If MousePointer = 11 Then Exit Sub
'  MousePointer = 11
'  If "" & pHundur.saekja("nr") = "" Then Exit Sub
'  pListar.porunarlisti pHundur.saekja("nr")
'  MousePointer = 0
End Sub

Private Sub Command4_Click()
  Load Form5
  'Form5.Grunnur pHundur.faGrunn
  Form5.hundur pHundur
  Form5.Stillingar pStillingar
  Form5.Show vbModal
  Birta
  Unload Form5
End Sub

Private Sub Command5_Click()
  Load Form8
  Form8.Listar pListar
  Form8.hundur pHundur
  Form8.Show vbModal
  Unload Form8
End Sub

Private Sub Command6_Click()
  Load Form3
  Form3.Grunnur pHundur.faGrunn
  Form3.hundur 0
  Form3.Stillingar pStillingar
  Form3.Show vbModal
  pHundur.setjaHund Form3.faHund
  Birta
  Unload Form3
End Sub

Private Sub Command7_Click()
  Load Form4
  Form4.hundur pHundur
  Form4.Stillingar pStillingar
  Form4.Show vbModal
  Birta
  Unload Form4
End Sub

Private Sub Command8_Click()
  Load Form5
  Form5.hundur pHundur
  Form5.Stillingar pStillingar
  Form5.Listar pListar
  Form5.porunarbok
  Form5.Show vbModal
  Birta
  Unload Form5
End Sub

Private Sub Command9_Click()
  Load Form10
  Form10.Grunnur pHundur.faGrunn
  Form10.Eigandi
  Form10.Show vbModal
  Unload Form10
End Sub

Private Sub Form_activate()
  If Not m_hladid Then
    Load Form9
    Form9.Stillingar pStillingar
    Form9.Show vbModal
    Form1.Caption = "CaniGree - " & Form9.cboTegund.Text & " - " & m_strUtg
    Unload Form9
    pHundur.Stillingar pStillingar
    pHundur.setjaHund
    pListar.Stillingar pStillingar
    pListar.Grunnur pHundur.faGrunn
    Birta
    If (dulkoda(pStillingar.Email, pStillingar.Leyfi) = False) Then
      lblLeyfi.Caption = "Þessi útgáfa er eingöngu til reynslu - fáðu leyfisnúmer á www.Canigree.com til að geta notað alla eiginleika forritsins."
    Else
      lblLeyfi.Caption = ""
    End If
    
  End If
  m_hladid = True
End Sub

Private Sub Form_Load()
  'Dim mynd As String
  Set pStillingar = New CStillingar
  Set pHundur = New CHundur
  Set pListar = New CListar
End Sub

Private Sub Birta()
  Dim mynd As String
  Dim ib As Integer, ih As Integer
  Dim id As IPictureDisp
  lblNafn.Caption = IIf(pHundur.saekja("titill") <> "", pHundur.saekja("titill") & " ", "") & pHundur.saekja("nafn") & vbCrLf
  If pHundur.saekja("titill") <> "" Then lblNafn.ForeColor = &HFF Else lblNafn.ForeColor = &H0
  
  If "" & pHundur.saekja("titill") <> "" Then txtStori.ForeColor = &HFF Else txtStori.ForeColor = &H0
  txtStori.Text = "Nafn: " & IIf(pHundur.saekja("titill") <> "", pHundur.saekja("titill") & " ", "") & pHundur.saekja("nafn") & vbCrLf
  txtStori.Text = txtStori.Text & IIf(pHundur.saekja("innkallsnafn") <> "", "Innkallsnafn: " & pHundur.saekja("innkallsnafn") & vbCrLf, "")
  txtStori.Text = txtStori.Text & "Fæðingardagur: " & pHundur.saekja("fdags") & vbCrLf
  txtStori.Text = txtStori.Text & "Ættbókarnúmer: " & pHundur.saekja("aettbokarnr") & vbCrLf
  txtStori.Text = txtStori.Text & IIf(pHundur.saekja("ormerki") <> "", "Örmerki: " & pHundur.saekja("ormerki") & vbCrLf, "")
  txtStori.Text = txtStori.Text & IIf(pHundur.saekja("litur") <> "", "Litur: " & pHundur.saekja("litur") & vbCrLf, "")
  txtStori.Text = txtStori.Text & "Kyn: " & IIf(pHundur.saekja("kyn") = "kk", "Karlhundur", IIf(pHundur.saekja("kyn") = "kvk", "Tík", "Ekki skráð")) & vbCrLf
  txtStori.Text = txtStori.Text & IIf(pHundur.saekja("landheiti") <> "", "Fæðingarstaður: " & pHundur.saekja("landheiti") & vbCrLf, "")
  txtStori.Text = txtStori.Text & IIf("" & pHundur.saekja("raektandinafn") <> "", "Ræktandi: " & pHundur.saekja("raektandinafn") & vbCrLf, "")
  txtStori.Text = txtStori.Text & IIf("" & pHundur.saekja("eigandinafn") <> "", "Eigandi: " & pHundur.saekja("eigandinafn") & vbCrLf, "")
  If pHundur.saekja("geldur") = "já" Then
    txtStori.Text = txtStori.Text & "Ekki til ræktunar: já" & vbCrLf
  End If
  If pHundur.saekja("latinn") = "já" Then
    txtStori.Text = txtStori.Text & IIf(pHundur.saekja("kyn") = "kvk", "Látin: ", "Látinn: ") & "já" & vbCrLf
  End If
  
  txtAthugasemd.Text = pHundur.saekja("athugasemd")
  
  lblFadir.Caption = IIf(pHundur.saekja("tf") <> "", pHundur.saekja("tf") & " ", "") & pHundur.saekja("nf")
  If pHundur.saekja("tf") <> "" Then lblFadir.ForeColor = &HFF Else lblFadir.ForeColor = &H0
  lblModir.Caption = IIf(pHundur.saekja("tm") <> "", pHundur.saekja("tm") & " ", "") & pHundur.saekja("nm")
  If pHundur.saekja("tm") <> "" Then lblModir.ForeColor = &HFF Else lblModir.ForeColor = &H0
  
  On Error Resume Next
  Image1 = LoadPicture
  Image1.Width = 2175
  Image1.Height = 2175
  mynd = pHundur.saekja("mynd1")
  If mynd <> "" Then
    Set id = LoadPicture(pStillingar.Mappa & pStillingar.Myndir & "\" & mynd)
    Image1 = id 'LoadPicture(pStillingar.Mappa & pStillingar.Myndir & "\" & mynd)
    ib = id.Width
    ih = id.Height
    Image1.Width = IIf(ib > ih, 2175, ib / ih * 2175)
    Image1.Height = IIf(ib > ih, ih / ib * 2175, 2175)
  End If
  
  mynd = pHundur.saekja("mynd2")
  Image2 = LoadPicture
  Image2.Width = 2175
  Image2.Height = 2175
  If mynd <> "" Then
    Set id = LoadPicture(pStillingar.Mappa & pStillingar.Myndir & "\" & mynd)
    Image2 = id 'LoadPicture(pStillingar.Mappa & pStillingar.Myndir & "\" & mynd)
    ib = id.Width
    ih = id.Height
    Image2.Width = IIf(ib > ih, 2175, ib / ih * 2175)
    Image2.Height = IIf(ib > ih, ih / ib * 2175, 2175)
    'Image2 = LoadPicture(pStillingar.Mappa & pStillingar.Myndir & "\" & mynd)
  End If
  On Error GoTo 0
  
  Dim nFyrir As Integer
  Dim i As Integer
  nFyrir = 0
  For i = 0 To lstSaga.ListCount - 1
    If pHundur.saekja("nr") <> "" Then
      If lstSaga.ItemData(i) = pHundur.saekja("nr") Then nFyrir = i + 1
    End If
  Next
  
  lstSaga.AddItem pHundur.saekja("nafn") & " - " & pHundur.saekja("FDags"), 0 'IIf(pHundur.saekja("titill") <> "", pHundur.saekja("titill") & " ", "") &
  lstSaga.ItemData(0) = "0" & pHundur.saekja("nr")
  If nFyrir > 0 Then
    lstSaga.RemoveItem nFyrir
  ElseIf lstSaga.ListCount > 10 Then
    lstSaga.RemoveItem 10
  End If
  lstSaga.ListIndex = 0
  
  BirtaAfkvaemi
  BirtaSystkini

  'DoEvents
  
  'lblIcGildi.Caption = Format(pListar.reiknaIC(pHundur.saekja("nr"), 0, 0), "0.0000%")
End Sub

Private Sub BirtaAfkvaemi()
  lstAfkvaemi.Clear
  
  If "" & pHundur.saekja("nr") = "" Then Exit Sub
  
  pHundur.setjaLista ("nrf=" & pHundur.saekja("nr") & " or nrm=" & pHundur.saekja("nr") & " and tegundid=" & pStillingar.Tegund & " order by fdags")
  Do
    lstAfkvaemi.AddItem IIf(pHundur.saekjaLista("latinn") = "fram", "* ", "") & pHundur.saekjaLista("nafn") & " - " & pHundur.saekjaLista("FDags") 'IIf(pHundur.saekjaLista("titill") <> "", pHundur.saekjaLista("titill") & " ", "") &
    lstAfkvaemi.ItemData(lstAfkvaemi.ListCount - 1) = "0" & pHundur.saekjaLista("nr")
  Loop Until pHundur.naestaLista
  lblAfkvaemi = "(" & pHundur.fjoldiLista & ")"
  pHundur.lokaLista
End Sub

Private Sub BirtaSystkini()
  Dim s As String
  lstSystkini.Clear
  
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

    lstSystkini.AddItem IIf(pHundur.saekjaLista("latinn") = "fram", "* ", "") & pHundur.saekjaLista("nafn") & " - " & pHundur.saekjaLista("FDags") & " " & s 'IIf(pHundur.saekjaLista("titill") <> "", pHundur.saekjaLista("titill") & " ", "") &
    lstSystkini.ItemData(lstSystkini.ListCount - 1) = "0" & pHundur.saekjaLista("nr")
  Loop Until pHundur.naestaLista
  lblSystkini = "(" & pHundur.fjoldiLista & ")"
  pHundur.lokaLista
End Sub

Private Sub cmdLeita_Click()
  Dim strLeita As String
  strLeita = Replace(txtLeita.Text, "'", "''")
  If optLeit1.Value = True Then
    pHundur.setjaLista ("(nafn like '*" & strLeita & "*' or fdags like '*" & strLeita & "*' or AettbokarNr like '*" & strLeita & "*' or ormerki like '*" & strLeita & "*' or titill like '*" & strLeita & "*') and tegundaflokkur='" & pStillingar.Tegundaflokkur & "' order by nafn, fdags")
  ElseIf optLeit2.Value = True Then
    pHundur.setjaLista ("(nafn like '*" & strLeita & "*' or fdags like '*" & strLeita & "*' or AettbokarNr like '*" & strLeita & "*' or ormerki like '*" & strLeita & "*' or titill like '*" & strLeita & "*') and klubbhundar='já' and tegundaflokkur='" & pStillingar.Tegundaflokkur & "' order by nafn, fdags")
  ElseIf optLeit3.Value = True Then
    pHundur.setjaLista ("(nafn like '*" & strLeita & "*' or fdags like '*" & strLeita & "*' or AettbokarNr like '*" & strLeita & "*' or ormerki like '*" & strLeita & "*' or titill like '*" & strLeita & "*') and eiginhundar='já' and tegundaflokkur='" & pStillingar.Tegundaflokkur & "' order by nafn, fdags")
  Else
    pHundur.setjaLista ("(nafn like '*" & strLeita & "*' or fdags like '*" & strLeita & "*' or AettbokarNr like '*" & strLeita & "*' or ormerki like '*" & strLeita & "*' or titill like '*" & strLeita & "*') and latinn='fram' and tegundaflokkur='" & pStillingar.Tegundaflokkur & "' order by nafn, fdags")
  End If
  lstLeit.Clear
  Do
    lstLeit.AddItem pHundur.saekjaLista("nafn") & " - " & pHundur.saekjaLista("FDags") 'IIf(pHundur.saekjaLista("titill") <> "", pHundur.saekjaLista("titill") & " ", "") &
    lstLeit.ItemData(lstLeit.ListCount - 1) = "0" & pHundur.saekjaLista("nr")
  Loop Until pHundur.naestaLista
  lblLeit = "(" & pHundur.fjoldiLista & ")"
  pHundur.lokaLista
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

Private Sub lstLeit_DblClick()
  pHundur.setjaHund lstLeit.ItemData(lstLeit.ListIndex)
  Birta
End Sub

Private Sub lstAfkvaemi_DblClick()
  pHundur.setjaHund lstAfkvaemi.ItemData(lstAfkvaemi.ListIndex)
  Birta
End Sub

Private Sub lstSaga_dblClick()
  pHundur.setjaHund lstSaga.ItemData(lstSaga.ListIndex)
  Birta
End Sub

Private Sub lstSystkini_DblClick()
  pHundur.setjaHund lstSystkini.ItemData(lstSystkini.ListIndex)
  Birta
End Sub

Private Sub lblFadir_DblClick()
  pHundur.setjaHund "0" & pHundur.saekja("nrf")
  Birta
End Sub

Private Sub lblModir_DblClick()
  pHundur.setjaHund "0" & pHundur.saekja("nrm")
  Birta
End Sub

Private Sub mnuHaetta_Click()
  End
End Sub

Private Sub mnuLesaInnCanigreeSkra_Click()
  Dim rs As Recordset
  Dim rs2 As Recordset
  Dim file
  Dim line As String
  Dim aline() As String
  Dim maxid As Integer
  
  If (dulkoda(pStillingar.Email, pStillingar.Leyfi) = False) Then
    MsgBox "Ekki er hægt að lesa inn í reynsluútgáfu forritsins.", vbOKOnly, "Innlestur"
    Exit Sub
  End If
  
  MousePointer = 11
  
  ' lesa inn í töfluna innlestur
  
  Set rs = pHundur.faGrunn.OpenRecordset("innlestur")
  While Not rs.EOF
    rs.Delete
    rs.MoveFirst
  Wend
  
  file = FreeFile
  
  'cdl.FileName = pStillingar.Mappa & "canigree " & Format(Now(), "yyyymmdd hhss") & ".dat"
  cdl.Filter = "Gagnaskrá|*.dat"
  cdl.CancelError = True
  'On Error GoTo loka:
  cdl.ShowOpen
  
  Open cdl.FileTitle For Input As file
  'Open pStillingar.Mappa & "canigree-" & pStillingar.Tegund & ".txt" For Input As file
  While Not EOF(file)
    Line Input #file, line
    rs.AddNew
    aline = Split(line, vbTab)
    rs!id = aline(0)
    rs!Nafn = aline(1)
    rs!titill = aline(2)
    rs!aettbokarnr = aline(3)
    Dim d As Date
    If aline(4) <> "" Then
      d = CDate(aline(4))
      rs!fdags = d
    Else
      rs!fdags = Null
    End If
    rs!kyn = aline(5)
    rs!tegundid = aline(6)
    rs!mynd1 = aline(7)
    rs!mynd2 = aline(8)
    rs!ormerki = aline(9)
    rs!innkallsnafn = aline(10)
    rs!litur = aline(11)
    rs!idf = 0 & aline(12)
    rs!idm = 0 & aline(13)
    rs!landid = 0 & aline(14)
    rs!klubbhundar = aline(15)
    rs!latinn = aline(16)
    rs!geldur = aline(17)
    rs!Raektandi = aline(18)
    rs!Eigandi = aline(19)
    rs.Update
    DoEvents
  Wend
  
  rs.MoveFirst
  
  While Not rs.EOF
    Dim fundid As Boolean
    fundid = False
    
    ' tengja ættbókarnúmer saman
    
    If rs!aettbokarnr <> "" Then
    rs.Edit
      Set rs2 = pHundur.faGrunn.OpenRecordset("select nr,fadirnr,modirnr from hundar where aettbokarnr='" & rs!aettbokarnr & "'")
      If Not rs2.EOF Then
        rs2.MoveLast
        If rs2.RecordCount = 1 Then
          rs!nr = rs2!nr
          rs!fadirnr = rs2!fadirnr
          rs!modirnr = rs2!modirnr
          fundid = True
        End If
      End If
      rs.Update
    End If
    
    ' tengja nöfn saman þar sem aðeins eitt nafn passar við innlestur
    
    If Not fundid And rs!Nafn <> "" Then
      rs.Edit
      Set rs2 = pHundur.faGrunn.OpenRecordset("select nr,fadirnr,modirnr from hundar where nafn=""" & rs!Nafn & """")
      If Not rs2.EOF Then
        rs2.MoveLast
        If rs2.RecordCount = 1 Then
          rs!nr = rs2!nr
          rs!fadirnr = rs2!fadirnr
          rs!modirnr = rs2!modirnr
          fundid = True
        ElseIf rs2.RecordCount > 1 Then
          rs!nr = -1
        End If
      End If
      rs.Update
    End If
    
    DoEvents
    rs.MoveNext
  Wend
  
  ' finna foreldra þeirra sem eru ótengdir
  
  Set rs = pHundur.faGrunn.OpenRecordset("select * from innlestur where fadirnr=0 or modirnr=0")
  While Not rs.EOF
    Set rs2 = pHundur.faGrunn.OpenRecordset("select nr from innlestur where id=" & rs!idf)
    If Not rs2.EOF Then
      rs.Edit
      rs!fadirnr = rs2!nr
      rs.Update
    End If
    rs2.Close
    
    Set rs2 = pHundur.faGrunn.OpenRecordset("select nr from innlestur where id=" & rs!idm)
    If Not rs2.EOF Then
      rs.Edit
      rs!modirnr = rs2!nr
      rs.Update
    End If
    rs2.Close
    
    ' tengja nafn og foreldra við innlestur
    
    Set rs2 = pHundur.faGrunn.OpenRecordset("select nr,fadirnr,modirnr from hundar where nafn=""" & rs!Nafn & """ and fadirnr=" & rs!fadirnr & " and modirnr=" & rs!modirnr)
    If Not rs2.EOF Then
      rs2.MoveLast
      If rs2.RecordCount = 1 Then
        rs.Edit
        rs!nr = rs2!nr
        rs.Update
        'rs!fadirnr = rs2!fadirnr
        'rs!modirnr = rs2!modirnr
        fundid = True
      'ElseIf rs2.RecordCount > 1 Then
      '  rs!nr = -1
      End If
    End If
    'rs.Update
    
    rs.MoveNext
  Wend
  rs.Close
  
  ' uppfæra nýjar foreldratengingar úr innlestri
  
  ' #### vantar
  
  ' úthluta númerum til nýrra einstaklinga og bæta við hundalistann
  
  pHundur.faGrunn.Execute ("delete * from innlestur where nr<>0")
    
  Set rs = pHundur.faGrunn.OpenRecordset("select max(nr) as maxnr from hundar")
  If Not rs.EOF Then
    maxid = 1
    On Error Resume Next
    maxid = rs!maxnr + 1
    On Error GoTo 0
  Else
    maxid = 1
  End If
  
  Set rs = pHundur.faGrunn.OpenRecordset("select * from innlestur where nr=0")
  While Not rs.EOF
    rs.Edit
    rs!nr = maxid
    
    rs.Update
    rs.MoveNext
    maxid = maxid + 1
  Wend
  
  ' tengja foreldra við innlestur
  
  Set rs = pHundur.faGrunn.OpenRecordset("select * from innlestur")
  While Not rs.EOF
    rs.Edit
    If rs!fadirnr = 0 Then
      Set rs2 = pHundur.faGrunn.OpenRecordset("select nr from innlestur where id=" & rs!idf)
      If Not rs2.EOF Then
        rs!fadirnr = rs2!nr
      End If
    End If
    If rs!modirnr = 0 Then
      Set rs2 = pHundur.faGrunn.OpenRecordset("select nr from innlestur where id=" & rs!idm)
      If Not rs2.EOF Then
        rs!modirnr = rs2!nr
      End If
    End If

    rs.Update
    rs.MoveNext
  Wend
  
  ' flytja gögn á milli yfir í hundalista

  Set rs = pHundur.faGrunn.OpenRecordset("select * from innlestur")
  While Not rs.EOF
    Set rs2 = pHundur.faGrunn.OpenRecordset("hundar")
    rs2.AddNew
    rs2!nr = rs!nr
    rs2!fadirnr = rs!fadirnr
    rs2!modirnr = rs!modirnr
    rs2!Nafn = rs!Nafn
    rs2!titill = rs!titill
    rs2!aettbokarnr = rs!aettbokarnr
    rs2!fdags = rs!fdags
    rs2!kyn = rs!kyn
    rs2!tegundid = rs!tegundid
    rs2!mynd1 = ""
    rs2!mynd2 = ""
    
    If rs!mynd1 <> "" And Dir(UCase(Left(cdl.FileName, Len(cdl.FileName) - Len(cdl.FileTitle))) & "\canigree_photos\" & rs!mynd1, vbNormal) <> "" Then
      FileCopy UCase(Left(cdl.FileName, Len(cdl.FileName) - Len(cdl.FileTitle))) & "\canigree_photos\" & rs!mynd1, pStillingar.Mappa & pStillingar.Myndir & "\" & rs!mynd1
      rs2!mynd1 = rs!mynd1
    End If
    
    If rs!mynd2 <> "" And Dir(UCase(Left(cdl.FileName, Len(cdl.FileName) - Len(cdl.FileTitle))) & "\canigree_photos\" & rs!mynd2, vbNormal) <> "" Then
      FileCopy UCase(Left(cdl.FileName, Len(cdl.FileName) - Len(cdl.FileTitle))) & "\canigree_photos\" & rs!mynd2, pStillingar.Mappa & pStillingar.Myndir & "\" & rs!mynd2
      rs2!mynd2 = rs!mynd2
    End If
    
    rs2!ormerki = rs!ormerki
    rs2!innkallsnafn = rs!innkallsnafn
    rs2!litur = rs!litur
    rs2!landid = rs!landid
    rs2!klubbhundar = rs!klubbhundar
    rs2!latinn = rs!latinn
    rs2!geldur = rs!geldur
    rs2!Raektandi = rs!Raektandi
    rs2!Eigandi = rs!Eigandi
    rs2.Update
    rs2.Close
    rs.MoveNext
    maxid = maxid + 1
  Wend
  
  pHundur.faGrunn.Execute ("delete * from innlestur")
  
  rs.Close
  Close file
  MsgBox "CaniGree skrá lesin inn", vbOKOnly, "Lesa inn CaniGree skrá"
loka:
  MousePointer = 0
End Sub

Private Sub mnuListarAettartre_Click()
Command5_Click
End Sub

Private Sub mnuListarAettbaekur_Click()
Command2_Click
End Sub

Private Sub mnuSkodaAettbok_Click()
Command7_Click
End Sub

Private Sub mnuSkodaICforfedur_Click()
  MousePointer = 11
  MsgBox pHundur.saekja("nafn") & vbCrLf & "IC-forfeður: " & vbCrLf & pListar.forfedurIC(0, pHundur.saekja("nrf"), pHundur.saekja("nrm"), vbCrLf), vbOKOnly, "IC-forfeður"
  MousePointer = 0
End Sub

Private Sub mnuSkodaICgildi_Click()
  MousePointer = 11
  MsgBox pHundur.saekja("nafn") & vbCrLf & "IC-gildi: " & Format(pListar.reiknaIC(0, pHundur.saekja("nrf"), pHundur.saekja("nrm")), "0.0000%"), vbOKOnly, "IC-gildi"
  MousePointer = 0
End Sub

Private Sub mnuSkodaPorun_Click()
Command8_Click
End Sub

Private Sub mnuSkodaPorunarlisti_Click()
Command3_Click
End Sub

Private Sub mnuSkraningBreytaSkraningu_Click()
Command1_Click
End Sub

Private Sub mnuSkraningEigendur_Click()
Command9_Click
End Sub

Private Sub mnuSkraningNyrHundur_Click()
Command6_Click
End Sub

Private Sub mnuSkraningSkraGot_Click()
Command4_Click
End Sub

Private Sub mnuSkraningVistaPED_Click()
  MousePointer = 11
  Dim rs As Recordset
  Dim file
  file = FreeFile
  Open pStillingar.Mappa & "hundar-" & pStillingar.Tegund & ".ped" For Output As file
  
  Print #file, "ID" & vbTab & "SireID" & vbTab & "DamID" & vbTab & "Nafn"
  Set rs = pHundur.faGrunn.OpenRecordset("select * from icgildi")
  While Not rs.EOF
    Print #file, rs!id & vbTab & rs!sireID & vbTab & rs!damID & vbTab & rs!Nafn
      DoEvents
      rs.MoveNext
  Wend
  
  rs.Close
  Close file
  MousePointer = 0
  MsgBox ".ped skráin er vistuð", vbOKOnly, "Vista .ped skrá"
End Sub

Private Sub mnuSkraningVistaCanigree_Click()
  MousePointer = 11
  Dim file
  Dim rs As Recordset
  file = FreeFile
  
  cdl.FileName = pStillingar.Mappa & "canigree " & Format(Now(), "yyyymmdd hhmm") & ".dat"
  cdl.Filter = "Gagnaskrá|*.dat"
  cdl.CancelError = True
  On Error GoTo loka:
  cdl.ShowSave
  On Error GoTo 0
  
  Open cdl.FileTitle For Output As file
  
  Set rs = pHundur.faGrunn.OpenRecordset("select * from hundalisti")
  While Not rs.EOF
    Print #file, rs!nr & vbTab & rs!Nafn & vbTab & rs!titill & vbTab & rs!aettbokarnr & vbTab & _
        Format(rs!fdags, "yyyy-mm-dd") & vbTab & rs!kyn & vbTab & rs!tegundid & vbTab & rs!mynd1 & vbTab & rs!mynd2 & vbTab _
      & rs!ormerki& & vbTab & rs!innkallsnafn & vbTab & rs!litur & vbTab & rs!fadirnr & vbTab & rs!modirnr _
      & vbTab & rs!landid & vbTab & rs!klubbhundar & vbTab & rs!latinn & vbTab & rs!geldur & vbTab & rs!Raektandi & vbTab & rs!Eigandi

    If Dir(UCase(Left(cdl.FileName, Len(cdl.FileName) - Len(cdl.FileTitle))) & "\canigree_photos\", vbDirectory) = "" Then
      MkDir (UCase(Left(cdl.FileName, Len(cdl.FileName) - Len(cdl.FileTitle))) & "\canigree_photos\")
    End If
    
    If rs!mynd1 <> "" And Dir(pStillingar.Mappa & pStillingar.Myndir & "\" & rs!mynd1, vbNormal) <> "" Then
      FileCopy pStillingar.Mappa & pStillingar.Myndir & "\" & rs!mynd1, UCase(Left(cdl.FileName, Len(cdl.FileName) - Len(cdl.FileTitle))) & "\canigree_photos\" & rs!mynd1
      
    End If
      
    If rs!mynd2 <> "" And Dir(pStillingar.Mappa & pStillingar.Myndir & "\" & rs!mynd2, vbNormal) <> "" Then
      FileCopy pStillingar.Mappa & pStillingar.Myndir & "\" & rs!mynd2, UCase(Left(cdl.FileName, Len(cdl.FileName) - Len(cdl.FileTitle))) & "\canigree_photos\" & rs!mynd2
    End If
    
    DoEvents
    rs.MoveNext
  Wend
  
  rs.Close
  Close file
  MsgBox "CaniGree skráin er vistuð", vbOKOnly, "Vista CaniGree skrá"
loka:
  MousePointer = 0
End Sub

Private Sub mnuVefurSkodaHund_Click()
  On Error Resume Next
  pListar.VistaAettbok pHundur.saekja("nr"), True
  On Error GoTo 0
'  Shell "explorer " & pStillingar.Mappa & pHundur.saekja("nr") & ".htm", vbNormalFocus
End Sub

Private Sub mnuVefurHundalisti_Click()
  pListar.VistaLista "deildar", True
'  Shell "explorer " & pStillingar.Mappa & "index-" & pStillingar.Tegund & ".htm", vbNormalFocus
End Sub

Private Sub txtAthugasemd_LostFocus()
Dim rs As Recordset
Set rs = pHundur.faGrunn.OpenRecordset("select * from hundar where nr=" & pHundur.saekja("nr"))
  rs.Edit
  rs.Fields("athugasemd") = txtAthugasemd.Text
  rs.Update
  rs.Close
End Sub

Private Sub txtLeita_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then cmdLeita_Click
End Sub

Private Sub txtStori_DblClick()
  Load Form4
  Form4.hundur pHundur
  Form4.Stillingar pStillingar
  Form4.Show vbModal
  Birta
  Unload Form4
End Sub
