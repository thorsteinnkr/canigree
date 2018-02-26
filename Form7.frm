VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form7 
   BackColor       =   &H00CDEBEB&
   Caption         =   "Myndir"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   LinkTopic       =   "Form7"
   ScaleHeight     =   4215
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdl 
      Left            =   2760
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00AFCDCD&
      Caption         =   "Leita á disk"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton cmdLoka 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "Loka"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdVistaHund 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "Vista"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDEBEB&
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4815
      Begin VB.ListBox lstSkrar 
         Height          =   2205
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblMyndir 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4575
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_db As Database
Dim m_skrarNafn As String
Dim pStillingar As CStillingar

Public Sub Grunnur(fdb As Database)
  Set m_db = fdb
End Sub

Public Sub Stillingar(fStillingar As CStillingar)
  Set pStillingar = fStillingar
  cdl.InitDir = pStillingar.Mappa & pStillingar.Myndir
End Sub

Private Sub Command1_Click()
  cdl.Filter = "Myndir|*.jpg;*.gif;*.bmp;*.png"
  cdl.CancelError = True
  On Error GoTo loka:
  cdl.ShowOpen
  If cdl.FileName <> pStillingar.Mappa & pStillingar.Myndir & "\" & cdl.FileTitle Then
    FileCopy cdl.FileName, pStillingar.Mappa & pStillingar.Myndir & "\" & cdl.FileTitle
  End If
  m_skrarNafn = cdl.FileTitle
  Me.Hide
loka:
End Sub

Private Sub Form_activate()
  Command1.Caption = "Leita á diski"
  Birta
End Sub

Sub Birta()
  Dim skra
  Dim rs As Recordset
  Dim i As Long
  lstSkrar.Clear
  lblMyndir.Caption = pStillingar.Mappa & pStillingar.Myndir
  skra = Dir(pStillingar.Mappa & pStillingar.Myndir & "\")
  While skra <> ""
    lstSkrar.AddItem skra
    skra = Dir
  Wend
  Set rs = m_db.OpenRecordset("select mynd1 from hundalisti where mynd1<>''")
  While Not rs.EOF
    For i = 0 To lstSkrar.ListCount - 1
      If UCase(lstSkrar.List(i)) = UCase(rs.Fields("mynd1")) Then lstSkrar.RemoveItem (i)
    Next
    rs.MoveNext
  Wend
  rs.Close
  Set rs = m_db.OpenRecordset("select mynd2 from hundalisti where mynd2<>''")
  While Not rs.EOF
    For i = 0 To lstSkrar.ListCount - 1
      If UCase(lstSkrar.List(i)) = UCase(rs.Fields("mynd2")) Then lstSkrar.RemoveItem (i)
    Next
    rs.MoveNext
  Wend
  rs.Close
End Sub

Public Function SkrarNafn() As String
  SkrarNafn = m_skrarNafn
End Function

Private Sub cmdLoka_Click()
  m_skrarNafn = ""
  Me.Hide
End Sub

Private Sub cmdVistaHund_Click()
  m_skrarNafn = lstSkrar.Text
  Me.Hide
End Sub


