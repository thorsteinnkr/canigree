VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStillingar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sMappa As String
Private m_sVolume As String
Private m_sMyndir As String
Private m_sGrunnur As String
Private m_iTegund As Long
Private m_sTegundheiti As String
Private m_strNafn As String
Private m_strEmail As String
Private m_strLeyfi As String
Private m_strSerial As String
Private m_iTegundaflokkur As Long


Private Sub Class_Initialize()
    Dim str As String
    Dim n As Integer
  
    Dim txt As String
    Dim volume_name As String * 256
    Dim file_system_name As String * 256
    Dim serial_number As Long
    Dim component_length As Long
    Dim system_flags As Long
    'Dim sectors_per_cluster As Long
    'Dim bytes_per_sector As Long
    'Dim free_clusters As Long
    'Dim total_clusters As Long
    'Dim total_bytes As Long
    'Dim free_bytes As Long

    If GetVolumeInformation(Left(App.Path, 3), _
      volume_name, Len(volume_name), _
      serial_number, component_length, _
      system_flags, file_system_name, _
      Len(file_system_name)) <> 0 _
    Then
      m_sVolume = Format$(serial_number)
    Else
      m_sVolume = ""
    End If
  
  str = App.Path
  m_strNafn = ""
  For n = 1 To Len(str)
    m_strNafn = m_strNafn & IIf(UCase(Mid(str, n, 1)) >= "A" And UCase(Mid(str, n, 1)) <= "Z", Mid(str, n, 1), "")
  Next
  m_strNafn = m_strNafn & m_sVolume
  m_sMappa = GetIniSetting("hundar", "stillingar_" & m_strNafn, "mappa", Split(App.Path, "\")(0) & "\CaniGree files\")
  'm_sMappa = Split(App.Path, "\")(0) & "\CaniGree files\"
  If Dir(m_sMappa, vbDirectory) = "" Then MkDir (m_sMappa)
  'm_sMyndir = GetIniSetting("hundar", "stillingar_" & m_strNafn, "myndir", "Photos")
  m_sMyndir = "Photos"
  If Dir(m_sMappa & m_sMyndir, vbDirectory) = "" Then MkDir (m_sMappa & m_sMyndir)
  
  m_sGrunnur = GetIniSetting("hundar", "stillingar_" & m_strNafn, "grunnur", App.Path & "\" & Dir(App.Path & "/*.mdb"))
  'm_sGrunnur = App.Path & "\..\data\CaniGree.mdb"
  
  m_iTegund = GetIniSetting("hundar", "stillingar_" & m_strNafn, "tegund", 1)
  m_sTegundheiti = GetIniSetting("hundar", "stillingar_" & m_strNafn, "tegundheiti", 1)

  m_strEmail = GetIniSetting("hundar", "stillingar_" & m_strNafn, "email", "")
  m_strLeyfi = GetIniSetting("hundar", "stillingar_" & m_strNafn, "leyfi", "")
  m_strSerial = GetIniSetting("hundar", "stillingar_" & m_strNafn, "password", "")
End Sub

'***
Property Get Mappa() As String
  Mappa = m_sMappa
End Property

Public Property Let Mappa(ByVal vNewValue As String)
  If Right(vNewValue, 1) <> "\" Then vNewValue = vNewValue & "\"
  m_sMappa = vNewValue
  SaveIniSetting "hundar", "stillingar_" & m_strNafn, "mappa", m_sMappa
End Property

'***
Property Get Myndir() As String
  Myndir = m_sMyndir
End Property

Public Property Let Myndir(ByVal vNewValue As String)
  If Right(vNewValue, 1) = "\" Then vNewValue = Left(vNewValue, Len(vNewValue) - 1)
  If Right(vNewValue, 1) = "/" Then vNewValue = Left(vNewValue, Len(vNewValue) - 1)
  m_sMyndir = vNewValue
  SaveIniSetting "hundar", "stillingar_" & m_strNafn, "myndir", m_sMyndir
End Property

'***
Public Property Get Grunnur() As String
  Grunnur = m_sGrunnur
End Property

Public Property Let Grunnur(ByVal vNewValue As String)
  m_sGrunnur = vNewValue
  SaveIniSetting "hundar", "stillingar_" & m_strNafn, "grunnur", m_sGrunnur
End Property

'***
Property Get Tegund() As Long
  Tegund = m_iTegund
End Property

Public Property Let Tegund(ByVal vNewValue As Long)
  m_iTegund = vNewValue
  SaveIniSetting "hundar", "stillingar_" & m_strNafn, "tegund", "" & m_iTegund
End Property

'***
Public Property Get Tegundheiti() As Variant
  Tegundheiti = m_sTegundheiti
End Property

Public Property Let Tegundheiti(ByVal vNewValue As Variant)
  m_sTegundheiti = vNewValue
End Property

'***
Public Property Get Email() As Variant
  Email = m_strEmail
End Property

Public Property Let Email(ByVal vNewValue As Variant)
  m_strEmail = vNewValue
  SaveIniSetting "hundar", "stillingar_" & m_strNafn, "email", "" & m_strEmail
End Property

'***
Public Property Get Leyfi() As Variant
  Leyfi = m_strLeyfi
End Property

Public Property Let Leyfi(ByVal vNewValue As Variant)
  m_strLeyfi = vNewValue
  SaveIniSetting "hundar", "stillingar_" & m_strNafn, "leyfi", "" & m_strLeyfi
End Property

'***
Public Property Get Serial() As Variant
  Serial = m_strSerial
End Property

Public Property Let Serial(ByVal vNewValue As Variant)
  m_strSerial = vNewValue
  SaveIniSetting "hundar", "stillingar_" & m_strNafn, "password", "" & m_strSerial
End Property

'***
Public Property Get Tegundaflokkur() As Variant
  Tegundaflokkur = m_iTegundaflokkur
End Property

Public Property Let Tegundaflokkur(ByVal vNewValue As Variant)
  m_iTegundaflokkur = vNewValue
End Property

Public Sub ReCreate()

End Sub

