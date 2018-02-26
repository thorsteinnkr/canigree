VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00CDEBEB&
   Caption         =   "Ættartré"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2685
   LinkTopic       =   "Form8"
   ScaleHeight     =   1965
   ScaleWidth      =   2685
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Sýna endurtekningar"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton cmdLoka 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "Loka"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdVistaHund 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "Vista"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Text            =   "5"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Dýpt ættartrés:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pListar As CListar
Private pHundur As CHundur
Private pStillingar As CStillingar

Dim bBuinn() As Boolean
Dim sListi(10000) As String
Dim nListi(10000) As Long
Dim iTeljari As Long
Dim file

Public Sub hundur(fHundur As CHundur)
  Set pHundur = fHundur
End Sub

Public Sub Listar(fListar As CListar)
  Set pListar = fListar
End Sub

Public Sub Stillingar(fStillingar As CStillingar)
  Set pStillingar = fStillingar
End Sub

Private Sub cmdVistaHund_Click()
  MousePointer = 11
  If "" & pHundur.saekja("nr") = "" Then Exit Sub
  pListar.Aettartre pHundur.saekja("nr"), Text1.Text, Check1.Value
  MousePointer = 0
  cmdLoka_Click
End Sub

Private Sub cmdLoka_Click()
  Me.Hide
End Sub

Private Sub Text1_Change()
  If Val(Text1.Text) > 12 Then
    Check1.Value = 0
    Check1.Enabled = False
  Else
    Check1.Enabled = True
  End If
End Sub
