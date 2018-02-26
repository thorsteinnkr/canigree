VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00CDEBEB&
   Caption         =   "Skrifa ættbækur"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3135
   LinkTopic       =   "Form6"
   ScaleHeight     =   5280
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkGotlisti 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Gotlisti"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Width           =   2655
   End
   Begin VB.CheckBox chkHundur 
      BackColor       =   &H00CDEBEB&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   2655
   End
   Begin VB.CheckBox chkListiMinRaektun 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Listi mín ræktun"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CheckBox chkListiAdrir 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Listi aðrir hundar"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CheckBox chkListiDeildar 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Listi deildarhundar"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CheckBox chkAdrir 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Aðrir hundar"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CheckBox chkDeildar 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Deildarhundar"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CheckBox chkMinRaektun 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Mín ræktun"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CheckBox chkEigin 
      BackColor       =   &H00CDEBEB&
      Caption         =   "Eigin hundar"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdLoka 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "Loka"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdVistaHund 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCDCD&
      Caption         =   "Vista"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Skrifa út yfirlitslista:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Skrifa út ættbók fyrir:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2895
   End
   Begin VB.Shape Shape3 
      Height          =   1575
      Left            =   120
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      Height          =   1935
      Left            =   120
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pListar As CListar
Dim pHundur As CHundur

Sub hundur(fHundur As CHundur)
  Set pHundur = fHundur
End Sub

Sub Listar(fListar As CListar)
  Set pListar = fListar
End Sub

Private Sub cmdVistaHund_Click()
  MousePointer = 11
  If chkHundur.Value = 1 Then pListar.VistaAettbok pHundur.saekja("nr"), True
  
  If chkEigin.Value = 1 Then pListar.VistaAettbaekur ("eigin")
  If chkMinRaektun.Value = 1 Then pListar.VistaAettbaekur ("mínræktun")
  If chkDeildar.Value = 1 Then pListar.VistaAettbaekur ("deildar")
  If chkAdrir.Value = 1 Then pListar.VistaAettbaekur ("aðrir")
  
  If chkListiMinRaektun.Value = 1 Then pListar.VistaLista ("mínræktun")
  If chkListiDeildar.Value = 1 Then pListar.VistaLista ("deildar")
  If chkListiAdrir.Value = 1 Then pListar.VistaLista ("aðrir")
  
  If chkGotlisti.Value = 1 Then pListar.Gotlisti
  
  MousePointer = 0
  If chkHundur.Value <> 1 And chkGotlisti.Value <> 1 Then
    MsgBox "Búið að vista ættbækurnar", vbOKOnly, "Ættbækur"
  End If
  cmdLoka_Click
End Sub

Private Sub cmdLoka_Click()
  Me.Hide
End Sub

Private Sub Form_activate()
  chkHundur.Caption = pHundur.saekja("nafn")
End Sub

