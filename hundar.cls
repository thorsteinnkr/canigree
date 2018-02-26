VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHundur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_db As Database
Dim m_bTengdur As Boolean
Dim m_lFyrstiHundur As Long
Dim m_rs As Recordset
Dim m_rsSystkin As Recordset
Dim m_rsHvolpar As Recordset
Dim m_rsListi As Recordset
Dim pStillingar As CStillingar

Private Sub Class_Initialize()
  m_lFyrstiHundur = GetIniSetting("hundar", "stillingar", "fyrstihundur", 1)
End Sub

Public Sub Stillingar(fStillingar As CStillingar)
  Set pStillingar = fStillingar
End Sub

Sub tengjaGrunn()
  If m_bTengdur = False Then
    Set m_db = OpenDatabase(pStillingar.Grunnur)
    m_bTengdur = True
  End If
End Sub

Function faGrunn() As Database
  tengjaGrunn
  Set faGrunn = m_db
End Function

Function fyrstiHundur() As Long
  fyrstiHundur = m_lFyrstiHundur
End Function

Sub setjaHund(Optional nr As Integer)
  tengjaGrunn
  If Not (m_rs Is Nothing) Then m_rs.Close
  If nr = 0 Then
    nr = fyrstiHundur
  Else
    m_lFyrstiHundur = nr
  End If
  Set m_rs = m_db.OpenRecordset("select * from hundalisti where nr=" & nr & " and tegundaflokkur='" & pStillingar.Tegundaflokkur & "'")
  If m_rs.RecordCount = 0 Then
    Dim max As Long
    Set m_rs = m_db.OpenRecordset("select max(nr) as maxnr from hundar")
    nr = IIf(IsNull(m_rs.Fields("maxnr")), 1, m_rs.Fields("maxnr"))
    m_rs.Close
    m_lFyrstiHundur = nr
    Set m_rs = m_db.OpenRecordset("select * from hundalisti where nr=" & nr & " and tegundaflokkur='" & pStillingar.Tegundaflokkur & "'")
  End If
  Set m_rsHvolpar = Nothing
  Set m_rsSystkin = Nothing
  SaveIniSetting "hundar", "stillingar", "fyrstihundur", "" & m_lFyrstiHundur
End Sub

Function saekja(svaedi As String) As String
  If Not m_rs.EOF Then
    saekja = "" & m_rs.Fields(svaedi)
  Else
    saekja = ""
  End If
End Function

'*****************************************************************
' Listar
'
'

Sub setjaLista(valid As String)
  tengjaGrunn
  Set m_rsListi = m_db.OpenRecordset("select * from hundalisti where " & valid)
End Sub

Function saekjaLista(svaedi As String) As String
  If m_rsListi Is Nothing Then setjaLista "*"
  If Not m_rsListi.EOF Then
    saekjaLista = "" & m_rsListi.Fields(svaedi)
  Else
    saekjaLista = ""
  End If
End Function

Function fjoldiLista() As Integer
  fjoldiLista = m_rsListi.RecordCount
End Function

Function naestaLista() As Boolean
  If Not (m_rsListi Is Nothing) And Not m_rsListi.EOF Then m_rsListi.MoveNext
  naestaLista = m_rsListi.EOF
End Function

Sub lokaLista()
  m_rsListi.Close
End Sub
