VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00CDEBEB&
   Caption         =   "Pörunarlisti"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2685
   LinkTopic       =   "Form11"
   ScaleHeight     =   2520
   ScaleWidth      =   2685
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check3 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Eingöngu framræktun"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Sýna framræktun"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CommandButton cmdVistaHund 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "Vista"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoka 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "Loka"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Sýna ættbækur hvolpa"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   2415
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pHundur As CHundur
Dim pListar As CListar

Public Sub hundur(fHundur As CHundur)
  Set pHundur = fHundur
End Sub

Sub Listar(fListar As CListar)
  Set pListar = fListar
End Sub

Private Sub Check2_Click()
  Check3.Enabled = IIf(Check2.Value = 1, True, False)
End Sub

Private Sub cmdLoka_Click()
  Me.Hide
End Sub

Private Sub cmdVistaHund_Click()
  If MousePointer = 11 Then Exit Sub
  MousePointer = 11
  If "" & pHundur.saekja("nr") = "" Then Exit Sub
  pListar.porunarlisti pHundur.saekja("nr"), IIf(Check1.Value = 1, True, False), IIf(Check2.Value = 1, True, False), IIf(Check3.Value = 1, True, False)
  MousePointer = 0
  cmdLoka_Click
End Sub

