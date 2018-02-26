VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CListar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_db As Database
Dim bBuinn() As Boolean
'Dim sListi(10000) As String
'Dim nListi(10000) As Long
Dim Listi(10000) As tListi
Dim iTeljari As Long
Dim file
Dim pStillingar As CStillingar
Dim m_hundaIC() As Double
Dim bHundaIC As Boolean
Dim m_hundar() As tHundar
Dim m_countHundar As Long

Private Type tListi
  fjoldi As Long
  nr As Long
  titill As String
  Nafn As String
  fdags As String
  aettbokarnr As String
  dypt As Long
  mynd1 As String
  mynd2 As String
  endurtekinn As Boolean
End Type

Private Type tIC
  lidur(50) As Long
  fjoldi As Long
  afkvaemi(50) As Long
End Type

Private Type tHundar
  sireID As Long
  damID As Long
  ICgildi As Double
End Type

Public Sub Stillingar(fStillingar As CStillingar)
  Set pStillingar = fStillingar
End Sub

Sub Grunnur(fdb As Database)
  Set m_db = fdb
End Sub

'*************************************************************
' Listar
'

Public Sub VistaLista(fGerd As String, Optional bSyna = False)
  Dim rs As Recordset
  Dim fj As Integer
  Dim file
  Dim sFDags As String
  
  If fGerd = "mínræktun" Then
    file = FreeFile
    Open pStillingar.Mappa & "minraektun-" & pStillingar.Tegund & ".htm" For Output As file
    Print #file, "<html><head><title>" & pStillingar.Tegundheiti & "</title></head><body>"
    Print #file, "<font size=5><center>" & pStillingar.Tegundheiti & "</center></font></p>"
    Print #file, "<hr size=3 noshade color=#000000><center><table>"

    Set rs = m_db.OpenRecordset("select * from hundalisti where minraektun='já' and latinn<>'fram' and tegundid=" & pStillingar.Tegund & " order by fdags, nr")
    While Not rs.EOF
      If sFDags <> rs!fdags And Not rs.BOF Then
        Print #file, "<tr><td>&nbsp;<td><td></tr>"
        sFDags = rs!fdags
      End If
      Print #file, "<tr bgcolor=#FFFFCC><td><a href=" & rs!nr & ".htm>"
      If "" & rs!titill <> "" Then
        Print #file, "<font color=red>" & rs!titill
      End If
      Print #file, rs!Nafn
      If "" & rs!titill <> "" Then
        Print #file, "</font>"
      End If
      Print #file, "</a><td>" & rs!fdags
      If rs!latinn = "já" Then
        Print "†"
      End If
      Print "<td><img src=" & rs!land & ".gif></tr>"
      
      fj = fj + 1
      'frmAettbaekur.lblTeljari = "Fjöldi: " & fj
      DoEvents
      rs.MoveNext
    Wend
    rs.Close
    
    Print #file, "</table>"
    Print #file, "</center><hr size=3 noshade color=#000000>"
    Print #file, "Höfundarréttur © 2000-" & Year(Now()) & " CaniGree.net <br>"
    Print #file, "<font size=2>" & Now() & "</font><p>"
    Print #file, "</body></html>"
    
    Close file
  End If
  
  
  If fGerd = "deildar" Then
    file = FreeFile
    Open pStillingar.Mappa & "index-" & pStillingar.Tegund & ".htm" For Output As file
    Print #file, "<html><head><title>" & pStillingar.Tegundheiti & "</title></head><body>"
    Print #file, "<font size=5><center>" & pStillingar.Tegundheiti & " deildarhundar.</center></font></p>"
    Print #file, "<table width=100% cellspacing='0'>"

    Set rs = m_db.OpenRecordset("select * from hundalisti where klubbhundar='já' and latinn<>'fram' and tegundid=" & pStillingar.Tegund & " order by fdags desc, nafn")
    While Not rs.EOF
      If sFDags <> "" & rs!fdags Then ' And Not rs.BOF Then
        Print #file, "</table><hr size=3 noshade color=#000000><table width=100% cellspacing='0'>" '<tr><td>&nbsp;<td><td></tr>"
        sFDags = "" & rs!fdags
      End If
      
      Print #file, "<tr bgcolor=#FFFFCC><td width='150'><td><a href=" & rs!nr & ".htm>"
      Print #file, IIf("" & rs!titill <> "", "<font color=red>" & rs!titill, "")
      Print #file, rs!Nafn
      If rs!mynd1 <> "" Or rs!mynd2 <> "" Then
        Print #file, " <img src='mynd.gif' border='0'>"
      End If
      Print #file, IIf("" & rs!titill <> "", "</font>", "")
        
      'End If
      
      Print #file, "</a><td width=100>" & rs!aettbokarnr & "<td width=100>" & rs!fdags
      If rs!latinn = "já" Then
        Print #file, "†"
      End If
      Print #file, "</tr>" '<img src=" & rs!land & ".gif>
      
      fj = fj + 1
      'frmAettbaekur.lblTeljari = "Fjöldi: " & fj
      DoEvents
      rs.MoveNext
    Wend
    rs.Close
    
    Print #file, "</table><hr size=3 noshade color=#000000><center><a href=adrir-" & pStillingar.Tegund & ".htm>Aðrir " & pStillingar.Tegundheiti & " hundar.</a>"
    Print #file, "</center><hr size=3 noshade color=#000000>"
    Print #file, "Höfundarréttur © 2000-" & Year(Now()) & " CaniGree.net<br>"
    Print #file, "<font size=2>" & Now() & "</font><p>"
    Print #file, "</body></html>"
    
    Close file
    
    If bSyna Then
      Shell "explorer " & pStillingar.Mappa & "index-" & pStillingar.Tegund & ".htm", vbNormalFocus
    End If
  End If
  
  If fGerd = "aðrir" Then
    file = FreeFile
    Open pStillingar.Mappa & "adrir-" & pStillingar.Tegund & ".htm" For Output As file
    Print #file, "<html><head><title>" & pStillingar.Tegundheiti & "</title></head><body>"
    Print #file, "<font size=5><center>Aðrir " & pStillingar.Tegundheiti & " hundar.</center></font></p>"
    Print #file, "<center><hr size=3 noshade color=#000000><table cellspacing='0'>"

    Set rs = m_db.OpenRecordset("select * from hundalisti where klubbhundar<>'já' and latinn<>'fram' and tegundid=" & pStillingar.Tegund & " order by nafn")
    While Not rs.EOF
      Print #file, "<tr bgcolor=#FFFFCC><td width=150><td><a href=" & rs!nr & ".htm>"
      If "" & rs!titill <> "" Then
        Print #file, "<font color=red>" & rs!titill
      End If
      Print #file, rs!Nafn
      If rs!mynd1 <> "" Or rs!mynd2 <> "" Then
        Print #file, " <img src='mynd.gif' border='0'>"
      End If
      If "" & rs!titill <> "" Then
        Print #file, "</font>"
      End If
      Print #file, "</a><td width=100>" & rs!aettbokarnr & "<td width=100>" & rs!fdags & "</tr>"
      
      fj = fj + 1
      'frmAettbaekur.lblTeljari = "Fjöldi: " & fj
      DoEvents
      rs.MoveNext
    Wend
    rs.Close
    
    Print #file, "</table><hr size=3 noshade color=#000000><a href=index-" & pStillingar.Tegund & ".htm>" & pStillingar.Tegundheiti & " deildarhundar.</a>"
    Print #file, "</center><hr size=3 noshade color=#000000>"
    Print #file, "Höfundarréttur © 2000-" & Year(Now()) & " CaniGree.net<br>"
    Print #file, "<font size=2>" & Now() & "</font><p>"
    Print #file, "</body></html>"
    
    Close file
  End If
  
  fj = 0
End Sub

'*************************************************************
' Ættbækur
'

Public Sub VistaAettbaekur(fGerd As String)
  Dim rs As Recordset
  Dim fj As Integer
  
  fj = 0
  
    If fGerd = "eigin" Then
      Set rs = m_db.OpenRecordset("select nr from hundalisti where eiginhundar = 'já' and latinn<>'fram' and tegundid=" & pStillingar.Tegund & " order by nr")
      While Not rs.EOF
        VistaAettbok rs!nr
        fj = fj + 1
        DoEvents
        rs.MoveNext
      Wend
      rs.Close
    End If
    
    If fGerd = "mínræktun" Then
      Set rs = m_db.OpenRecordset("select nr from hundalisti where minraektun='já' and latinn<>'fram' and tegundid=" & pStillingar.Tegund & " order by nr")
      While Not rs.EOF
        VistaAettbok rs!nr
        fj = fj + 1
        DoEvents
        rs.MoveNext
      Wend
      rs.Close
    End If
    
    If fGerd = "deildar" Then
      Set rs = m_db.OpenRecordset("select nr from hundalisti where klubbhundar='já' and latinn<>'fram' and tegundid=" & pStillingar.Tegund & " order by nr")
      While Not rs.EOF
        VistaAettbok rs!nr
        fj = fj + 1
        DoEvents
        rs.MoveNext
      Wend
      rs.Close
    End If
 
  If fGerd = "aðrir" Then
    Set rs = m_db.OpenRecordset("select nr from hundalisti where eiginhundar<>'já' and minraektun<>'já' and klubbhundar<>'já' and latinn<>'fram' and tegundid=" & pStillingar.Tegund)
    While Not rs.EOF
      VistaAettbok rs!nr
        fj = fj + 1
      DoEvents
      rs.MoveNext
    Wend
    rs.Close
  End If

End Sub

Public Sub VistaAettbok(nr As Integer, Optional bSyna As Boolean)
  Dim rs As Recordset
  Dim rs2 As Recordset
  Dim file
  Dim im As IPictureDisp
  Dim ih As Integer
  Dim strHaed As String
  
  file = FreeFile

  Set rs = m_db.OpenRecordset("select * from hundalisti where nr=" & nr & " and tegundid=" & pStillingar.Tegund)
  If rs.EOF Then Exit Sub
    Open pStillingar.Mappa & nr & ".htm" For Output As file
    Print #file, "<html><head><title>" & rs!titill & " " & rs!Nafn & "</title>"
    Print #file, "<link href='canigree.css' rel='stylesheet' type='text/css'></head>"
    If rs!kyn = "kk" Then
      Print #file, "<body class='kk'>" 'bgcolor=#CFDFEF
    Else
      Print #file, "<body class='kvk'>" 'bgcolor=#FFCFCF
    End If
    Print #file, "<p class='heading1'>" & rs!titill & " " & rs!Nafn & "</p>"
    Print #file, "<hr>" 'size=3 noshade color=#000000
    Print #file, "" '<center>
    Print #file, "<p class='toppics'>"
    Print #file, "<table border='0'><tr>"
    If sNoNulls(rs!mynd1) <> "" Then
      strHaed = ""
      On Error Resume Next
      Set im = LoadPicture(pStillingar.Mappa & pStillingar.Myndir & "/" & rs!mynd1)
      ih = im.Height
      On Error GoTo 0
      strHaed = IIf(ih > 10583, "height='400'", "")
      Print #file, "<td><img src=""" & pStillingar.Myndir & "/" & rs!mynd1 & """ " & strHaed & "></td>"
    End If
    If sNoNulls(rs!mynd2) <> "" Then
      strHaed = ""
      On Error Resume Next
      Set im = LoadPicture(pStillingar.Mappa & pStillingar.Myndir & "/" & rs!mynd2)
      ih = im.Height
      On Error GoTo 0
      strHaed = IIf(ih > 10583, "height='400'", "")
      Print #file, "<td><img src=""" & pStillingar.Myndir & "/" & rs!mynd2 & """" & strHaed & "></td>"
    End If
    Print #file, "</tr>"
    Print #file, "</table>"
    Print #file, "</p>" '</center>

'almennar upplýsingar
    Print #file, "<table border=0 width=100%><tr>"
    Print #file, "<td width=50%><p>Titill: " & rs!titill & "</p></td>"
    Print #file, "<td width=50%><p>Ættbókarnúmer: " & rs!aettbokarnr & "</p></td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%><p>Innkallsnafn: " & rs!innkallsnafn & "</p></td>"
    Print #file, "<td width=50%><p>Fæðingardagur: " & IIf(rs!latinn = "fram", "<i>(Framræktun)</i>", rs!fdags & IIf(rs!latinn = "já", IIf(rs!kyn = "kk", " - Látinn", " - Látin"), "")) & "</p></td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%><p>Litur: " & rs!litur & "</p></td>"
    Print #file, "<td width=50%><p>Tegund: " & rs!Tegundheiti & "</p></td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%><p>Kyn:"
    If rs!kyn = "kk" Then
      Print #file, "hundur"
    ElseIf rs!kyn = "kvk" Then
      Print #file, "tík"
    End If
    Print #file, "</p></td>"
    Print #file, "<td width=50%><p>Ræktandi: " & rs!raektandinafn & "</p></td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%><p>"
    Print #file, "Eigandi: " & rs!eigandinafn & "</p></td>"
    Print #file, "<td width=50%><p>Fæðingarstaður: " & rs!landheiti & "</p></td>"
'    Print #file, "</tr><tr>"
'
'    Print #file, "<td width=50%>&nbsp;</td>"
'    Print #file, "<td width=50%>HD: " & rs!hd & "</td>"
    Print #file, "</tr></table>"
    
' Ættir
    Print #file, "<hr>" ' size=3 noshade color=#000000
    Print #file, "<p class='heading2'>Ættir:</p><p>"
    
    If rs!kyn = "kk" Then
      Print #file, "<table border=1 width=100% bordercolor=#CFDFEF bgcolor=#DFEAF4><tr>"
    Else
      Print #file, "<table border=1 width=100% bordercolor=#FFCFCF bgcolor=#FFDDDD><tr>"
    End If
    
    Print #file, listahund(rs, 4, "f")
    Print #file, listahund(rs, 2, "ff")
    Print #file, listahund(rs, 1, "fff")
    Print #file, listahund(rs, 1, "mff")
    Print #file, listahund(rs, 2, "mf")
    Print #file, listahund(rs, 1, "fmf")
    Print #file, listahund(rs, 1, "mmf")
    Print #file, listahund(rs, 4, "m")
    Print #file, listahund(rs, 2, "fm")
    Print #file, listahund(rs, 1, "ffm")
    Print #file, listahund(rs, 1, "mfm")
    Print #file, listahund(rs, 2, "mm")
    Print #file, listahund(rs, 1, "fmm")
    Print #file, listahund(rs, 0, "mmm")
    Print #file, "</tr></table>"
    
    If rs!latinn = "fram" Then
      Dim dIC As Double
      Dim strIC As String
      dIC = reiknaIC(0, rs!nrf, rs!nrm)
      strIC = IIf(dIC > 0.0625, "<font color=red>", "")
      strIC = strIC & IIf(dIC > 0.125, "<b>", "")
      strIC = strIC & "IC-Gildi: " & Format(dIC, "0.0000%")
      strIC = strIC & IIf(dIC > 0.125, "</b>", "")
      strIC = strIC & IIf(dIC > 0.0625, "</font>", "")
      Print #file, strIC
      Print #file, "<p>" & forfedurIC(0, rs!nrf, rs!nrm) & "</p>"
    End If
    
    Print #file, "</p><hr>" ' size=3 noshade color=#000000
    
'Hvolpar
    Set rs2 = m_db.OpenRecordset("select * from hundalisti where (fadirnr=" & nr & " or modirnr=" & nr & ") and latinn<>'fram' and tegundid=" & pStillingar.Tegund & " order by fdags,nafn")
    Print #file, "<p class='heading2'>Hvolpar: (" & rs2.RecordCount & ")</p>"
    Print #file, "<p class='offsprings'>"
'    Print #file, "<table border='0'>"
    If rs!kyn = "kk" Then
      Print #file, "<table border='1' bordercolor='#CFDFEF'><tr>" 'width=100%
    Else
      Print #file, "<table border='1' bordercolor='#FFCFCF'><tr>"
    End If
    While Not rs2.EOF
      Print #file, "<tr><td width='85'>"
      If rs2!mynd2 & "" <> "" Then
        Print #file, "<img src=""" & pStillingar.Myndir & "/" & rs2!mynd2 & """ width='80' align='middle'>"
      ElseIf rs2!mynd1 & "" <> "" Then
        Print #file, "<img src=""" & pStillingar.Myndir & "/" & rs2!mynd1 & """ width='80' align='middle'>"
      Else
        Print #file, "&nbsp;"
      End If
      Print #file, "<td><p><a href=" & rs2!nr & ".htm>"
      If sNoNulls(rs2!titill) <> "" Then Print #file, "<font color=red>" & rs2!titill & " "
      Print #file, rs2!Nafn
      If sNoNulls(rs2!titill) <> "" Then Print #file, "</font>"
      Print #file, "</a><br>" & IIf(rs2!aettbokarnr & "" <> "", rs2!aettbokarnr & "<br>", "") & rs2!fdags & "</p>"
      Print #file, "<td><p>" & IIf(rs2!nrf = rs!nr, "( x <a href=" & rs2!nrm & ".htm>" & IIf("" & rs2!tm <> "", rs2!tm & " ", "") & rs2!nm & "</a>)", "( x <a href=" & rs2!nrf & ".htm>" & IIf("" & rs2!tf <> "", rs2!tf & " ", "") & rs2!nf & "</a>)</p>")
      Print #file, "</tr>"
      rs2.MoveNext
    Wend
    rs2.Close
    Print #file, "</table></p>"
    Print #file, "<hr>" ' size=3 noshade color=#000000
    
'Systkin
    Set rs2 = m_db.OpenRecordset("select * from hundalisti where ((fadirnr=" & rs!fadirnr & " and fadirnr>0) or (modirnr=" & rs!modirnr & ") and modirnr>0) and nr<>" & nr & " and latinn<>'fram' and tegundid=" & pStillingar.Tegund & " order by fdags,nafn")
    Print #file, "<p class='heading2'>Systkin: (" & rs2.RecordCount & ")</p>"
    Print #file, "<p class='siblings'>"
'    Print #file, "<table border='0'>"
    If rs!kyn = "kk" Then
      Print #file, "<table border='1' bordercolor='#CFDFEF'><tr>" 'width=100%
    Else
      Print #file, "<table border='1' bordercolor='#FFCFCF'><tr>"
    End If
    While Not rs2.EOF
      Print #file, "<tr><td width='85'>"
      If rs2!mynd2 & "" <> "" Then
        Print #file, "<img src=""" & pStillingar.Myndir & "/" & rs2!mynd2 & """ width='80' align='middle'>"
      ElseIf rs2!mynd1 & "" <> "" Then
        Print #file, "<img src=""" & pStillingar.Myndir & "/" & rs2!mynd1 & """ width='80' align='middle'>"
      Else
        Print #file, "&nbsp;"
      End If
      Print #file, "<td><p><a href=" & rs2!nr & ".htm>"
      If sNoNulls(rs2!titill) <> "" Then Print #file, "<font color=red>" & rs2!titill & " "
      Print #file, rs2!Nafn
      If sNoNulls(rs2!titill) <> "" Then Print #file, "</font>"
      Print #file, "</a><br>" & IIf(rs2!aettbokarnr & "" <> "", rs2!aettbokarnr & "<br>", "") & rs2!fdags & "</p>"
'      Print #file, "</a><td><p>" & rs2!fdags & "</p>"
      If rs2!fadirnr = rs!fadirnr And rs2!modirnr = rs!modirnr Then
        Print #file, "<td><p>Alsystkin</p>"
      ElseIf rs2!fadirnr = rs!fadirnr Then
        Print #file, "<td><p>Samfeðra</p>"
      ElseIf rs2!modirnr = rs!modirnr Then
        Print #file, "<td><p>Sammæðra</p>"
      End If
      Print #file, "</tr>"
      rs2.MoveNext
    Wend
    rs2.Close
    Print #file, "</table></p>"
    Print #file, "<hr>" ' size=3 noshade color=#000000
    Print #file, "<p class='heading2'><a class='heading2' href='index-" & pStillingar.Tegund & ".htm'>Til baka í lista yfir " & pStillingar.Tegundheiti & " deildarhundana.</a></p>"
    Print #file, "<hr>" ' size=3 noshade color=#000000
    Print #file, "<p>Höfundarréttur © 2000-" & Year(Now()) & " CaniGree.net<br>"
    Print #file, "<i>" & Now() & "</i></p>"
    
    Print #file, "</body></html>"
    Close file
  rs.Close
  
  If bSyna Then Shell "explorer " & pStillingar.Mappa & nr & ".htm", vbNormalFocus

End Sub

'*************************************************************
' Pörunarbók
'

Public Sub porunarbok(hundur As Long, tik As Long, Optional bFram As Boolean)
  Dim rs As Recordset
  Dim rs2 As Recordset
  Dim rs3 As Recordset
  Dim file
  Dim strHaed As String
  Dim im As IPictureDisp
  Dim ih As Integer
  
  file = FreeFile

  Set rs = m_db.OpenRecordset("select * from hundalisti where nr=" & hundur & " and tegundid=" & pStillingar.Tegund)
  Set rs3 = m_db.OpenRecordset("select * from hundalisti where nr=" & tik & " and tegundid=" & pStillingar.Tegund)
  If rs.EOF Or rs3.EOF Then Exit Sub
    Open pStillingar.Mappa & hundur & "-" & tik & ".htm" For Output As file
    Print #file, "<html><head><title>" & rs!titill & " " & rs!Nafn & " - " & rs3!titill & " " & rs3!Nafn & "</title>"
    Print #file, "<link href='canigree.css' rel='stylesheet' type='text/css'></head>"
    Print #file, "</head>"
    Print #file, "<body bgcolor=#CFEFCF>"
    Print #file, "<p class='heading1'>" & rs!titill & " " & rs!Nafn & "</p>"
'    Print #file, "<hr size=3 noshade color=#000000>"
    Print #file, "<p class='toppics'>"
    Print #file, "<table border=0>"
    Print #file, "<tr>"
    If sNoNulls(rs!mynd1) <> "" Then
      strHaed = ""
      On Error Resume Next
      Set im = LoadPicture(pStillingar.Mappa & pStillingar.Myndir & "/" & rs!mynd1)
      ih = im.Height
      On Error GoTo 0
      strHaed = IIf(ih > 10583, "height='400'", "")
      Print #file, "<td><img src=""" & pStillingar.Myndir & "/" & rs!mynd1 & """ " & strHaed & "></td>"
'      Print #file, "<td><img src=" & pStillingar.Myndir & "/" & rs!mynd1 & "></td>"
    End If
    If sNoNulls(rs!mynd2) <> "" Then
      strHaed = ""
      On Error Resume Next
      Set im = LoadPicture(pStillingar.Mappa & pStillingar.Myndir & "/" & rs!mynd2)
      ih = im.Height
      On Error GoTo 0
      strHaed = IIf(ih > 10583, "height='400'", "")
      Print #file, "<td><img src=""" & pStillingar.Myndir & "/" & rs!mynd2 & """ " & strHaed & "></td>"
'      Print #file, "<td><img src=" & pStillingar.Myndir & "/" & rs!mynd2 & "></td>"
    End If
    Print #file, "</tr>"
    Print #file, "</table>"
    Print #file, "</p>"
'almennar upplýsingar
    Print #file, "<table border=0 width=100%><tr>"
    Print #file, "<td width=50%><p>Titill: " & rs!titill & "</p></td>"
    Print #file, "<td width=50%><p>Ættbókarnúmer: " & rs!aettbokarnr & "</p></td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%><p>Innkallsnafn: " & rs!innkallsnafn & "</p></td>"
'    Print #file, "<td width=50%><p>Fæðingardagur: " & rs!fdags & IIf(rs!latinn = "já", IIf(rs!kyn = "kk", " - Látinn", " - Látin"), "") & "</p></td>"
    Print #file, "<td width=50%><p>Fæðingardagur: " & IIf(rs!latinn = "fram", "<i>(Framræktun)</i>", rs!fdags & IIf(rs!latinn = "já", IIf(rs!kyn = "kk", " - Látinn", " - Látin"), "")) & "</p></td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%><p>Litur: " & rs!litur & "</p></td>"
    Print #file, "<td width=50%><p>Tegund: " & rs!Tegundheiti & "</p></td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%><p>Kyn:"
    If rs!kyn = "kk" Then
      Print #file, "hundur"
    ElseIf rs!kyn = "kvk" Then
      Print #file, "tík"
    End If
    Print #file, "</p></td>"
    Print #file, "<td width=50%><p>Ræktandi: " & rs!raektandinafn & "</p></td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%>"
    Print #file, "<p>Eigandi: " & rs!eigandinafn & "</p></td>"
    Print #file, "<td width=50%><p>Fæðingarstaður: " & rs!landheiti & "</p></td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%>&nbsp;</td>"
'    Print #file, "<td width=50%>HD: " & rs!hd & "</td>"
    Print #file, "</tr></table>"
'tíkin
    Print #file, "<p class='heading1'>" & rs3!titill & " " & rs3!Nafn & "</p>"
    Print #file, "<p class='toppics'>"
    Print #file, "<table border=0>"
    Print #file, "<tr>"
    If sNoNulls(rs3!mynd1) <> "" Then
      Print #file, "<td><img src=" & pStillingar.Myndir & "/" & rs3!mynd1 & "></td>"
    End If
    If sNoNulls(rs3!mynd2) <> "" Then
      Print #file, "<td><img src=" & pStillingar.Myndir & "/" & rs3!mynd2 & "></td>"
    End If
    Print #file, "</tr>"
    Print #file, "</table>"
    Print #file, "</p>"
'almennar upplýsingar
    Print #file, "<table border=0 width=100%><tr>"
    Print #file, "<td width=50%><p>Titill: " & rs3!titill & "</p></td>"
    Print #file, "<td width=50%><p>Ættbókarnúmer: " & rs3!aettbokarnr & "</p></td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%><p>Innkallsnafn: " & rs3!innkallsnafn & "</p></td>"
'    Print #file, "<td width=50%><p>Fæðingardagur: " & rs3!fdags & IIf(rs3!latinn = "já", IIf(rs3!kyn = "kk", " - Látinn", " - Látin"), "") & "</p></td>"
    Print #file, "<td width=50%><p>Fæðingardagur: " & IIf(rs3!latinn = "fram", "<i>(Framræktun)</i>", rs3!fdags & IIf(rs3!latinn = "já", IIf(rs3!kyn = "kk", " - Látinn", " - Látin"), "")) & "</p></td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%><p>Litur: " & rs3!litur & "</p></td>"
    Print #file, "<td width=50%><p>Tegund: " & rs3!Tegundheiti & "</p></td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%><p>Kyn:"
    If rs3!kyn = "kk" Then
      Print #file, "hundur"
    ElseIf rs3!kyn = "kvk" Then
      Print #file, "tík"
    End If
    Print #file, "</p></td>"
    Print #file, "<td width=50%><p>Ræktandi: " & rs3!raektandinafn & "</p></td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%>"
    Print #file, "<p>Eigandi: " & rs3!eigandinafn & "</p></td>"
    Print #file, "<td width=50%><p>Fæðingarstaður: " & rs3!landheiti & "</p></td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%>&nbsp;</td>"
'    Print #file, "<td width=50%>HD: " & rs!hd & "</td>"
    Print #file, "</tr></table>"
'
    Print #file, "<hr>"
    Print #file, "<p class='heading2'>Ættir:</p><p>"
    
    Print #file, "<table border=1 width=100% bordercolor=#CFEFCF bgcolor=#DFF4DF><tr>"
    Print #file, listahund2(rs!nr, rs!titill, rs!Nafn, rs!fdags, 8, 25)
    Print #file, listahund2(rs!nrf, rs!tf, rs!nf, rs!fdf, 4, 25)
    Print #file, listahund2(rs!nrff, rs!tff, rs!nff, rs!fdff, 2, 25)
    Print #file, listahund2(rs!nrfff, rs!tfff, rs!nfff, "", 1, 25)
    Print #file, listahund2(rs!nrmff, rs!tmff, rs!nmff, "", 1, 25)
    Print #file, listahund2(rs!nrmf, rs!tmf, rs!nmf, rs!fdmf, 2, 25)
    Print #file, listahund2(rs!nrfmf, rs!tfmf, rs!nfmf, "", 1, 25)
    Print #file, listahund2(rs!nrmmf, rs!tmmf, rs!nmmf, "", 1, 25)
    
    Print #file, listahund2(rs!nrm, rs!tm, rs!nm, rs!fdm, 4, 25)
    Print #file, listahund2(rs!nrfm, rs!tfm, rs!nfm, rs!fdfm, 2, 25)
    Print #file, listahund2(rs!nrffm, rs!tffm, rs!nffm, "", 1, 25)
    Print #file, listahund2(rs!nrmfm, rs!tmfm, rs!nmfm, "", 1, 25)
    Print #file, listahund2(rs!nrmm, rs!tmm, rs!nmm, rs!fdmm, 2, 25)
    Print #file, listahund2(rs!nrfmm, rs!tfmm, rs!nfmm, "", 1, 25)
    Print #file, listahund2(rs!nrmmm, rs!tmmm, rs!nmmm, "", 1, 25)
'    Print #file, "</tr></table></center>"
    
'    Print #file, "<table border=1 width=100%><tr>"
    Print #file, listahund2(rs3!nr, rs3!titill, rs3!Nafn, rs3!fdags, 8, 25)
    Print #file, listahund2(rs3!nrf, rs3!tf, rs3!nf, rs3!fdf, 4, 25)
    Print #file, listahund2(rs3!nrff, rs3!tff, rs3!nff, rs3!fdff, 2, 25)
    Print #file, listahund2(rs3!nrfff, rs3!tfff, rs3!nfff, "", 1, 25)
    Print #file, listahund2(rs3!nrmff, rs3!tmff, rs3!nmff, "", 1, 25)
    Print #file, listahund2(rs3!nrmf, rs3!tmf, rs3!nmf, rs3!fdmf, 2, 25)
    Print #file, listahund2(rs3!nrfmf, rs3!tfmf, rs3!nfmf, "", 1, 25)
    Print #file, listahund2(rs3!nrmmf, rs3!tmmf, rs3!nmmf, "", 1, 25)
    
    Print #file, listahund2(rs3!nrm, rs3!tm, rs3!nm, rs3!fdm, 4, 25)
    Print #file, listahund2(rs3!nrfm, rs3!tfm, rs3!nfm, rs3!fdfm, 2, 25)
    Print #file, listahund2(rs3!nrffm, rs3!tffm, rs3!nffm, "", 1, 25)
    Print #file, listahund2(rs3!nrmfm, rs3!tmfm, rs3!nmfm, "", 1, 25)
    Print #file, listahund2(rs3!nrmm, rs3!tmm, rs3!nmm, rs3!fdmm, 2, 25)
    Print #file, listahund2(rs3!nrfmm, rs3!tfmm, rs3!nfmm, "", 1, 25)
    Print #file, listahund2(rs3!nrmmm, rs3!tmmm, rs3!nmmm, "", 1, 25)
    Print #file, "</tr></table>"
    
    Dim dIC As Double
    Dim strIC As String
    
    dIC = reiknaIC(0, rs!nr, rs3!nr)
    If dIC <> -1 Then
        strIC = IIf(dIC > 0.0625, "<font color=red>", "")
        strIC = strIC & IIf(dIC > 0.125, "<b>", "")
        strIC = strIC & "IC-Gildi: " & Format(dIC, "0.0000%")
        strIC = strIC & IIf(dIC > 0.125, "</b>", "")
        strIC = strIC & IIf(dIC > 0.0625, "</font>", "")
        Print #file, strIC
    End If
    Print #file, "</p>"
    Print #file, "<p>" & forfedurIC(0, rs!nr, rs3!nr) & "</p>"
'    Print #file, "<p>IC-gildi: " & Format(reiknaIC(0, rs!nr, rs3!nr), "0.0000%") & "</p>"
    
    
    Print #file, "<hr>"
'    Set rs2 = m_db.OpenRecordset("select * from hundalisti where fadirnr=" & hundur & " or modirnr=" & hundur & " and tegundid=" & pStillingar.Tegund & " order by fdags,nafn")
'    Print #file, "<center>Hvolpar: (" & rs2.RecordCount & ")<p>"
'    Print #file, "<table>"
'    While Not rs2.EOF
'      Print #file, "<tr><td><a href=" & rs2!nr & ".htm>"
'      If sNoNulls(rs2!titill) <> "" Then Print #file, "<font color=red>" & rs2!titill & " "
'      Print #file, rs2!Nafn
'      If sNoNulls(rs2!titill) <> "" Then Print #file, "</font>"
'      Print #file, "</a><td>" & rs2!fdags & "</tr>"
'      rs2.MoveNext
'    Wend
'    rs2.Close
'    Print #file, "</table></center>"
'    Print #file, "<hr size=3 noshade color=#000000>"
    Set rs2 = m_db.OpenRecordset("select * from hundalisti where ((fadirnr=" & rs!nr & " and fadirnr>0) or (modirnr=" & rs3!nr & ") and modirnr>0) and latinn<>'fram' and tegundid=" & pStillingar.Tegund & " order by fdags,nafn")
    Print #file, "<p class='heading2'>Systkin: (" & rs2.RecordCount & ")</p>"
    Print #file, "<p class='siblings'>"
    Print #file, "<table>"
    While Not rs2.EOF
      Print #file, "<tr><td><a href=" & rs2!nr & ".htm>"
      If sNoNulls(rs2!titill) <> "" Then Print #file, "<font color=red>" & rs2!titill & " "
      Print #file, rs2!Nafn
      If sNoNulls(rs2!titill) <> "" Then Print #file, "</font>"
      Print #file, "</a><td><p>" & rs2!fdags & "</p></td>"
      If rs2!fadirnr = rs!nr And rs2!modirnr = rs3!nr Then
        Print #file, "<td><p>Alsystkin</p></td>"
      ElseIf rs2!fadirnr = rs!nr Then
        Print #file, "<td><p>Samfeðra</p></td>"
      ElseIf rs2!modirnr = rs3!nr Then
        Print #file, "<td><p>Sammæðra</p></td>"
      End If
      Print #file, "</tr>"
      rs2.MoveNext
    Wend
    rs2.Close
    Print #file, "</table></p>"
    
    If bFram = True Then
      
    End If
    
    Print #file, "<hr><p class='heading2'><a href=index-" & pStillingar.Tegund & ".htm>Til baka í lista yfir " & pStillingar.Tegundheiti & " deildarhundana.</a></p>"
    Print #file, "<hr><p>Höfundarréttur © 2000-" & Year(Now()) & " CaniGree.net<br>"
    Print #file, "<i>" & Now() & "</i></p>"
    
    Print #file, "</body></html>"
    Close file
  rs.Close
  rs3.Close
  
  Shell "explorer " & pStillingar.Mappa & hundur & "-" & tik & ".htm", vbNormalFocus
End Sub


'*************************************************************
' Pörunarlisti
'

Public Sub porunarlisti(hundur As Long, Optional bSynaAett As Boolean, Optional bSynaFram As Boolean, Optional bEingonguFram As Boolean)
  Dim file As Integer
  Dim rs As Recordset
  Dim rs2 As Recordset
  Dim rs3 As Recordset
  Dim rsK1 As Recordset
  Dim rsK2 As Recordset
  Dim sSkra As String
  Dim kyn As String
  Dim dIC As Double
  Dim systkin() As tHundar
  Dim systkin_fjoldi As Integer
  Dim i As Integer
  Dim j As Integer
  
  Set rs = m_db.OpenRecordset("select * from hundalisti where nr=" & hundur)
  If rs.EOF Then Exit Sub
  
  kyn = rs.Fields("kyn")
  
    Set rs2 = m_db.OpenRecordset("select fadirnr,modirnr, min(nr) as numer from hundalisti where (klubbhundar='já' " & IIf(bSynaFram = True, " or latinn='fram' ", "") & ") " & IIf(bEingonguFram = True, " and latinn='fram'", "") & " and kyn='" & IIf(kyn = "kk", "kvk", "kk") & "' and latinn<>'já' and geldur<>'já' and tegundaflokkur='" & pStillingar.Tegundaflokkur & "' group by fadirnr, modirnr")
  If rs2.EOF Then Exit Sub
  
  i = 0
  systkin_fjoldi = rs2.RecordCount
  ReDim systkin(systkin_fjoldi) As tHundar
  While Not rs2.EOF
    i = i + 1
    systkin(i).sireID = rs2!fadirnr
    systkin(i).damID = rs2!modirnr
    If kyn = "kk" Then
      systkin(i).ICgildi = reiknaIC(0, rs!nr, rs2!numer)
    Else
      systkin(i).ICgildi = reiknaIC(0, rs2!numer, rs!nr)
    End If
    rs2.MoveNext
  Wend
  rs2.Close
  
  Dim bHefRadad As Boolean
  Dim xS As Integer, xD As Integer, xI As Double
  bHefRadad = True
  i = 1
  While bHefRadad
    bHefRadad = False
    For j = 1 To systkin_fjoldi - 1
      If systkin(j + 1).ICgildi < systkin(j).ICgildi Then
        xS = systkin(j).sireID: xD = systkin(j).damID: xI = systkin(j).ICgildi
        systkin(j).sireID = systkin(j + 1).sireID: systkin(j).damID = systkin(j + 1).damID: systkin(j).ICgildi = systkin(j + 1).ICgildi
        systkin(j + 1).sireID = xS: systkin(j + 1).damID = xD: systkin(j + 1).ICgildi = xI
        bHefRadad = True
      End If
    Next
    i = i + 1
  Wend
 
  sSkra = pStillingar.Mappa & "s" & hundur & ".htm"
  
  file = FreeFile
  Open sSkra For Output As file

  Print #file, "<html><head><title>Pörunarlisti " & rs!titill & " " & rs!Nafn & "</title>"
  Print #file, "<link href='canigree.css' rel='stylesheet' type='text/css'></head>"
  Print #file, "</head>"
  If bSynaAett = True Then Print #file, "<STYLE>H3 {page-break-before: always}</STYLE>"
  Print #file, "<BODY><p class='heading1'>Pörunarlisti</p><hr>"

    Print #file, "<p class='heading1'>" & rs!titill & " " & rs!Nafn
    If DateDiff("yyyy", rs!fdags, Now()) >= 7 And rs!kyn = "kvk" Then Print #file, "¹"
    Print #file, "</p>"

    Print #file, "<center>"
    Print #file, "<table border=0>"
    Print #file, "<tr>"
    If sNoNulls(rs!mynd1) <> "" Then
      Print #file, "<td><img src=" & pStillingar.Myndir & "/" & rs!mynd1 & "></td>"
    End If
    If sNoNulls(rs!mynd2) <> "" Then
      Print #file, "<td><img src=" & pStillingar.Myndir & "/" & rs!mynd2 & "></td>"
    End If
    Print #file, "</tr>"
    Print #file, "</table>"
    Print #file, "</center>"
'almennar upplýsingar
    Print #file, "<table border=0 width=100%><tr>"
    Print #file, "<td width=50%><p>Titill: " & rs!titill & "</p></td>"
    Print #file, "<td width=50%><p>Ættbókarnúmer: " & rs!aettbokarnr & "</p></td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%><p>Innkallsnafn: " & rs!innkallsnafn & "</p></td>"
    Print #file, "<td width=50%><p>Fæðingardagur: " & rs!fdags & IIf(rs!latinn = "já", IIf(rs!kyn = "kk", " - Látinn", " - Látin"), "") & "</p></td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%><p>Litur: " & rs!litur & "</p></td>"
    Print #file, "<td width=50%><p>Tegund: " & rs!Tegundheiti & "</p></td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%><p>Kyn:"
    If rs!kyn = "kk" Then
      Print #file, "hundur"
    ElseIf rs!kyn = "kvk" Then
      Print #file, "tík"
    End If
    Print #file, "</p></td>"
    Print #file, "<td width=50%><p>Ræktandi: " & rs!raektandinafn & "</p></td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%>"
    Print #file, "<p>Eigandi: " & rs!eigandinafn & "</p></td>"
    Print #file, "<td width=50%><p>Fæðingarstaður: " & rs!landheiti & "</p></td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%>&nbsp;</td>"
'    Print #file, "<td width=50%>HD: " & rs!hd & "</td>"
    Print #file, "</tr></table>"
  
    i = 1
    While i <= systkin_fjoldi
      Dim iFnr As Long, iMnr As Long
      Dim str, strIC
      Dim bBreytt
      iFnr = systkin(i).sireID
      iMnr = systkin(i).damID
    
      str = "<p>"
      Set rs2 = m_db.OpenRecordset("select * from hundalisti where fadirnr=" & iFnr & " and modirnr=" & iMnr & " and (klubbhundar='já' or latinn='fram') and kyn='" & IIf(kyn = "kk", "kvk", "kk") & "' and latinn<>'já' and geldur<>'já' and tegundaflokkur='" & pStillingar.Tegundaflokkur & "' order by nafn")
      While Not rs2.EOF
        'MsgBox (DateDiff("yyyy", rs2!fdags, Now()))
      'Dim rs3 As Recordset
      Dim str3 As String
      Set rs3 = m_db.OpenRecordset("select * from hundar where nr=" & rs2.Fields("nr"))
      If Not rs3.EOF Then
        If rs3!mynd2 <> "" Then
          str = str & "<img src=""" & pStillingar.Myndir & "/" & rs3!mynd2 & """ width='80'><br>"
        ElseIf rs3!mynd1 <> "" Then
          str = str & "<img src=""" & pStillingar.Myndir & "/" & rs3!mynd1 & """ width='80'><br>"
        End If
      End If
      rs3.Close
  
        str = str & "<a href=" & rs2!nr & ".htm>"
        If sNoNulls(rs2!titill) <> "" Then str = str & "<font color=red>" & rs2!titill & " "
        str = str & rs2!Nafn
        If DateDiff("yyyy", rs2!fdags, Now()) >= 7 Then str = str & "¹" 'And rs2!kyn = "kvk"
        If sNoNulls(rs2!titill) <> "" Then str = str & "</font>"
        str = str & "</a>"
        If rs2!latinn = "fram" Then
          str = str & "<br><i>(Framræktun)</i><br>"
        Else
          str = str & "<br>" & rs2!fdags & "<br>"
        End If
        rs2.MoveNext
      Wend
      If Not rs2.BOF Then
        rs2.MovePrevious
      End If
      
      Print #file, "</p>"
        
      If kyn = "kvk" Then
        Set rsK1 = rs2
        Set rsK2 = rs
      Else
        Set rsK1 = rs
        Set rsK2 = rs2
      End If
        
      dIC = systkin(i).ICgildi
      strIC = IIf(dIC > 0.0625, "<font color=red>", "")
      strIC = strIC & IIf(dIC > 0.125, "<b>", "")
      strIC = strIC & "<p>IC-Gildi: " & Format(dIC, "0.0000%") & "</p>"
      strIC = strIC & IIf(dIC > 0.125, "</b>", "")
      strIC = strIC & IIf(dIC > 0.0625, "</font>", "")
        
      If bSynaAett = True And dIC < 0.25 Then
        Print #file, "<H3><hr size=3 noshade color=#000000></H3>"
        Print #file, "<table border=1 width=100%><tr>"
        If kyn = "kk" Then
          Print #file, listahund2(rsK1!nr, rsK1!titill, rsK1!Nafn, rsK1!fdags, 8, 25)
        Else
          Print #file, listahund2(0, "", str, "", 8, 25)
        End If
        Print #file, listahund(rsK1, 4, "f", , 25)
        Print #file, listahund(rsK1, 2, "ff", , 25)
        Print #file, listahund(rsK1, 1, "fff", False, 25)
        Print #file, listahund(rsK1, 1, "mff", False, 25)
        Print #file, listahund(rsK1, 2, "mf", , 25)
        Print #file, listahund(rsK1, 1, "fmf", False, 25)
        Print #file, listahund(rsK1, 1, "mmf", False, 25)
        Print #file, listahund(rsK1, 4, "m", , 25)
        Print #file, listahund(rsK1, 2, "fm", , 25)
        Print #file, listahund(rsK1, 1, "ffm", False, 25)
        Print #file, listahund(rsK1, 1, "mfm", False, 25)
        Print #file, listahund(rsK1, 2, "mm", , 25)
        Print #file, listahund(rsK1, 1, "fmm", False, 25)
        Print #file, listahund(rsK1, 1, "mmm", False, 25)
        If kyn = "kvk" Then
          Print #file, listahund2(rsK2!nr, rsK2!titill, rsK2!Nafn, rsK2!fdags, 8, 25)
        Else
          Print #file, listahund2(0, "", str, "", 8, 25)
        End If
        Print #file, listahund(rsK2, 4, "f", , 25)
        Print #file, listahund(rsK2, 2, "ff", , 25)
        Print #file, listahund(rsK2, 1, "fff", False, 25)
        Print #file, listahund(rsK2, 1, "mff", False, 25)
        Print #file, listahund(rsK2, 2, "mf", , 25)
        Print #file, listahund(rsK2, 1, "fmf", False, 25)
        Print #file, listahund(rsK2, 1, "mmf", False, 25)
        Print #file, listahund(rsK2, 4, "m", , 25)
        Print #file, listahund(rsK2, 2, "fm", , 25)
        Print #file, listahund(rsK2, 1, "ffm", False, 25)
        Print #file, listahund(rsK2, 1, "mfm", False, 25)
        Print #file, listahund(rsK2, 2, "mm", , 25)
        Print #file, listahund(rsK2, 1, "fmm", False, 25)
        Print #file, listahund(rsK2, 0, "mmm", False, 25)
        Print #file, "</tr></table>"
        Print #file, strIC
      ElseIf dIC < 0.25 Then
        Print #file, "<hr size=3 noshade color=#000000>"
        If kyn = "kk" Then
          Print #file, "<p>" & "(<a href=" & rsK2!nrf & ".htm>" & IIf("" & rsK2!tf <> "", rsK2!tf & " ", "") & rsK2!nf & "</a> x <a href=" & rsK2!nrm & ".htm>" & IIf("" & rsK2!tm <> "", rsK2!tm & " ", "") & rsK2!nm & "</a>)</p>"
        Else
          Print #file, "<p>" & "(<a href=" & rsK1!nrf & ".htm>" & IIf("" & rsK1!tf <> "", rsK1!tf & " ", "") & rsK1!nf & "</a> x <a href=" & rsK1!nrm & ".htm>" & IIf("" & rsK1!tm <> "", rsK1!tm & " ", "") & rsK1!nm & "</a>)</p>"
        End If
        Print #file, "<p>" & str & "</p>"
        Print #file, strIC
      End If
      
      i = i + 1
    Wend
    
    Print #file, "</table>"
    Print #file, "<hr>"
    Print #file, "<p>¹ Vegna aldurs gæti verið að ekki sé hægt að nota viðkomandi hund/tík til ræktunar.</p>"
    Print #file, "<hr>"
    Print #file, "<p>Höfundarréttur © 2000-" & Year(Now()) & " CaniGree.net<br>"
    Print #file, "<font size=2>" & Now() & "</font></p>"
    Print #file, "</body></html>"
  
  Close file
  Shell "explorer " & sSkra, vbNormalFocus

  rs.Close
  rs2.Close
End Sub

Public Sub porunarlisti2(hundur As Long, Optional bSynaAett As Boolean)
  Dim file As Integer
  Dim rs As Recordset
  Dim rs2 As Recordset
  Dim rs3 As Recordset
  Dim rsK1 As Recordset
  Dim rsK2 As Recordset
  Dim sSkra As String
  Dim kyn As String
  Dim dIC As Double
  
  Set rs = m_db.OpenRecordset("select * from hundalisti where nr=" & hundur)
  If rs.EOF Then Exit Sub
  
  kyn = rs.Fields("kyn")
  
  Set rs2 = m_db.OpenRecordset("select * from hundalisti where klubbhundar='já' and kyn='" & IIf(kyn = "kk", "kvk", "kk") & "' and latinn<>'já' and geldur<>'já' and tegundid=" & pStillingar.Tegund & " order by nf, nm")
'  Set rs2 = m_db.OpenRecordset("select * from hundalisti where klubbhundar='já' and kyn='" & IIf(kyn = "kk", "kvk", "kk") & "' and latinn<>'já' and geldur<>'já' and tegundid=" & pStillingar.Tegund & " and (fadirnr<>" & rs!fadirnr & " and modirnr<>" & rs!modirnr & ") order by nf, nm")
  If rs2.EOF Then Exit Sub
  
  sSkra = pStillingar.Mappa & "s" & hundur & ".htm"
  
  file = FreeFile
  Open sSkra For Output As file

  Print #file, "<html><head><title>Pörunarlisti " & rs!titill & " " & rs!Nafn & "</title></head>"
  If bSynaAett = True Then Print #file, "<STYLE>H3 {page-break-before: always}</STYLE>"
  Print #file, "<BODY><p><font size=5><center>Pörunarlisti</center></font></p><hr size=3 noshade color=#000000>"

    Print #file, "<p><font size=5><center>" & rs!titill & " " & rs!Nafn & "</center></font></p>"

    Print #file, "<center>"
    Print #file, "<table border=0>"
    Print #file, "<tr>"
    If sNoNulls(rs!mynd1) <> "" Then
      Print #file, "<td><img src=" & pStillingar.Myndir & "/" & rs!mynd1 & "></td>"
    End If
    If sNoNulls(rs!mynd2) <> "" Then
      Print #file, "<td><img src=" & pStillingar.Myndir & "/" & rs!mynd2 & "></td>"
    End If
    Print #file, "</tr>"
    Print #file, "</table>"
    Print #file, "</center><p>"
'almennar upplýsingar
    Print #file, "<table border=0 width=100%><tr>"
    Print #file, "<td width=50%>Titill: " & rs!titill & "</td>"
    Print #file, "<td width=50%>Ættbókarnúmer: " & rs!aettbokarnr & "</td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%>Innkallsnafn: " & rs!innkallsnafn & "</td>"
    Print #file, "<td width=50%>Fæðingardagur: " & rs!fdags & IIf(rs!latinn = "já", IIf(rs!kyn = "kk", " - Látinn", " - Látin"), "") & "</td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%>Litur: " & rs!litur & "</td>"
    Print #file, "<td width=50%>Tegund: " & rs!Tegundheiti & "</td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%>Kyn:"
    If rs!kyn = "kk" Then
      Print #file, "hundur"
    ElseIf rs!kyn = "kvk" Then
      Print #file, "tík"
    End If
    Print #file, "</td>"
    Print #file, "<td width=50%>Ræktandi: " & rs!raektandinafn & "</td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%>"
    Print #file, "Eigandi: " & rs!eigandinafn & "</td>"
    Print #file, "<td width=50%>Fæðingarstaður: " & rs!landheiti & "</td>"
    Print #file, "</tr><tr>"
    
    Print #file, "<td width=50%>&nbsp;</td>"
'    Print #file, "<td width=50%>HD: " & rs!hd & "</td>"
    Print #file, "</tr></table>"
  
    Dim iFnr As Long, iMnr As Long
    Dim str, strIC
    Dim bBreytt
    iFnr = rs2!fadirnr
    iMnr = rs2!modirnr
'    Print #file, "<H3><br></H3><table border=0 align='center'>"
    
    str = ""
    Do 'While Not rs2.EOF
      
      str = str & "<a href=" & rs2!nr & ".htm>"
      If sNoNulls(rs2!titill) <> "" Then str = str & "<font color=red>" & rs2!titill & " "
      str = str & rs2!Nafn
      If sNoNulls(rs2!titill) <> "" Then str = str & "</font>"
      str = str & "</a><br>" & rs2!fdags & "<br>"
      
      rs2.MoveNext
      
      If rs2.EOF Then
        bBreytt = True
        rs2.MovePrevious
      ElseIf iFnr <> rs2!fadirnr Or iMnr <> rs2!modirnr Then
        bBreytt = True
        rs2.MovePrevious
      Else
        bBreytt = False
      End If
      
      If bBreytt Then
        Print #file, "<br>" '<hr size=3 noshade color=#000000>"
'        Print #file, "</table><hr size=3 noshade color=#000000>"
        
        If kyn = "kvk" Then
          Set rsK1 = rs2
          Set rsK2 = rs
        Else
          Set rsK1 = rs
          Set rsK2 = rs2
        End If
        
        dIC = reiknaIC(0, rsK1!nr, rsK2!nr)
        strIC = IIf(dIC > 0.0625, "<font color=red>", "")
        strIC = strIC & IIf(dIC > 0.125, "<b>", "")
        strIC = strIC & "IC-Gildi: " & Format(dIC, "0.0000%")
        strIC = strIC & IIf(dIC > 0.125, "</b>", "")
        strIC = strIC & IIf(dIC > 0.0625, "</font>", "")
        
        If bSynaAett = True And dIC < 0.25 Then
          Print #file, "<H3><hr size=3 noshade color=#000000></H3>"
          Print #file, "<table border=1 width=100%><tr>"
          If kyn = "kk" Then
            Print #file, listahund2(rsK1!nr, rsK1!titill, rsK1!Nafn, rsK1!fdags, 8, 25)
          Else
            Print #file, listahund2(0, "", str, "", 8, 25)
          End If
          Print #file, listahund(rsK1, 4, "f", , 25)
          Print #file, listahund(rsK1, 2, "ff", , 25)
          Print #file, listahund(rsK1, 1, "fff", False, 25)
          Print #file, listahund(rsK1, 1, "mff", False, 25)
          Print #file, listahund(rsK1, 2, "mf", , 25)
          Print #file, listahund(rsK1, 1, "fmf", False, 25)
          Print #file, listahund(rsK1, 1, "mmf", False, 25)
          Print #file, listahund(rsK1, 4, "m", , 25)
          Print #file, listahund(rsK1, 2, "fm", , 25)
          Print #file, listahund(rsK1, 1, "ffm", False, 25)
          Print #file, listahund(rsK1, 1, "mfm", False, 25)
          Print #file, listahund(rsK1, 2, "mm", , 25)
          Print #file, listahund(rsK1, 1, "fmm", False, 25)
          Print #file, listahund(rsK1, 1, "mmm", False, 25)
          If kyn = "kvk" Then
            Print #file, listahund2(rsK2!nr, rsK2!titill, rsK2!Nafn, rsK2!fdags, 8, 25)
          Else
            Print #file, listahund2(0, "", str, "", 8, 25)
          End If
          Print #file, listahund(rsK2, 4, "f", , 25)
          Print #file, listahund(rsK2, 2, "ff", , 25)
          Print #file, listahund(rsK2, 1, "fff", False, 25)
          Print #file, listahund(rsK2, 1, "mff", False, 25)
          Print #file, listahund(rsK2, 2, "mf", , 25)
          Print #file, listahund(rsK2, 1, "fmf", False, 25)
          Print #file, listahund(rsK2, 1, "mmf", False, 25)
          Print #file, listahund(rsK2, 4, "m", , 25)
          Print #file, listahund(rsK2, 2, "fm", , 25)
          Print #file, listahund(rsK2, 1, "ffm", False, 25)
          Print #file, listahund(rsK2, 1, "mfm", False, 25)
          Print #file, listahund(rsK2, 2, "mm", , 25)
          Print #file, listahund(rsK2, 1, "fmm", False, 25)
          Print #file, listahund(rsK2, 0, "mmm", False, 25)
          Print #file, "</tr></table>"
          Print #file, strIC
        ElseIf dIC < 0.25 Then
          Print #file, "<hr size=3 noshade color=#000000>"
          If kyn = "kk" Then
            Print #file, "<p>" & "(<a href=" & rsK2!nrf & ".htm>" & IIf("" & rsK2!tf <> "", rsK2!tf & " ", "") & rsK2!nf & "</a> x <a href=" & rsK2!nrm & ".htm>" & IIf("" & rsK2!tm <> "", rsK2!tm & " ", "") & rsK2!nm & "</a>)</p>"
          Else
            Print #file, "<p>" & "(<a href=" & rsK1!nrf & ".htm>" & IIf("" & rsK1!tf <> "", rsK1!tf & " ", "") & rsK1!nf & "</a> x <a href=" & rsK1!nrm & ".htm>" & IIf("" & rsK1!tm <> "", rsK1!tm & " ", "") & rsK1!nm & "</a>)</p>"
          End If
          Print #file, "<p>" & str & "</p>"
          Print #file, strIC
        End If
        
        
        rs2.MoveNext
        If Not rs2.EOF Then
          iFnr = rs2!fadirnr
          iMnr = rs2!modirnr
        End If
        str = ""
'        Print #file, "<table border=0 align='center'>"
      End If
      
    Loop Until rs2.EOF 'Wend
    
    Print #file, "</table>"
    Print #file, "<hr size=3 noshade color=#000000>"
    Print #file, "Höfundarréttur © 2000-" & Year(Now()) & " CaniGree.net<br>"
    Print #file, "<font size=2>" & Now() & "</font><p>"
    Print #file, "</body></html>"
  
  Close file
  Shell "explorer " & sSkra, vbNormalFocus

  rs.Close
  rs2.Close
End Sub

Sub Gotlisti()
  Dim sSkra As String
  Dim rs As Recordset
  Dim fyrraar As String
  sSkra = pStillingar.Mappa & "gotlisti-" & pStillingar.Tegund & ".htm"
  file = FreeFile
  Open sSkra For Output As file
  Set rs = m_db.OpenRecordset("select * from gotlisti where tegundid=" & pStillingar.Tegund & " order by fdags desc")
  Print #file, "<html><head><title>Gotlisti</title></head>"
  Print #file, "<BODY><p><font size=5><center>Gotlisti</center></font></p>"
  
  Print #file, "<table  width=100% cellspacing=0>"
  While Not rs.EOF
  On Error Resume Next
    If Not rs.BOF And Year(rs!fdags & "") <> fyrraar Then Print #file, "</tr></table><hr size=3 noshade color=#000000><table width=100% cellspacing=0>"
    On Error GoTo 0
    Print #file, "<tr bgcolor=#FFFFCC><td width=30%>" & "<a href=" & rs!nrf & ".htm>" & IIf("" & rs!tf <> "", "<font color=red>" & rs!tf & " ", "") & rs!nf & IIf("" & rs!tf <> "", "</font>", "") & "</a>"
    Print #file, " <td  width=30%> x " & "<a href=" & rs!nrm & ".htm>" & IIf("" & rs!tm <> "", "<font color=red>" & rs!tm & " ", "") & rs!nm & IIf("" & rs!tm <> "", "</font>", "") & "</a>"
    Print #file, " <td  width=10%> " & rs!fdags & " <td  width=10%> " & rs!fjoldi & IIf(rs!fjoldi > 1, " hvolpar <td  width=20%>", " hvolpur <td>")
    
    Dim dIC As Double
    Dim strIC As String
    On Error Resume Next
    dIC = reiknaIC(0, rs!nrf, rs!nrm)
    strIC = IIf(dIC > 0.0625, "<font color=red>", "")
    strIC = strIC & IIf(dIC > 0.125, "<b>", "")
    strIC = strIC & " IC: " & Format(dIC, "0.0000%")
    strIC = strIC & IIf(dIC > 0.125, "</b>", "")
    strIC = strIC & IIf(dIC > 0.0625, "</font>", "")
    On Error GoTo 0
    Print #file, strIC
    
    Print #file, "</tr>"
    On Error Resume Next
    fyrraar = Year(rs!fdags)
    On Error GoTo 0
    rs.MoveNext
  Wend
  Print #file, "</table>"
  
  Print #file, "<hr size=3 noshade color=#000000>"
  Print #file, "Höfundarréttur © 2000-" & Year(Now()) & " CaniGree.net<br>"
  Print #file, "<font size=2>" & Now() & "</font><p>"
  Print #file, "</body></html>"
  
  Close file
  Shell "explorer " & sSkra, vbNormalFocus

  rs.Close
End Sub

'*************************************************************
' Ættartré
'

Sub Aettartre(numer As Integer, maxdypi As Integer, Endur As Integer)
  Dim rs As Recordset
  Dim max As Long
  Dim i As Long
  Set rs = m_db.OpenRecordset("select max(nr) as maxnr from hundar")
  max = rs.Fields("maxnr")
  rs.Close
    
  file = FreeFile
  Open pStillingar.Mappa & "t" & numer & ".htm" For Output As file
  
  iTeljari = 0
  ReDim bBuinn(max) As Boolean
  forHundur numer, 0, maxdypi, Endur
  
  Print #file, "<html><head><title>Ættartré - " & Listi(1).Nafn & "</title>"
  Print #file, "<link href='canigree.css' rel='stylesheet' type='text/css'></head>"
  Print #file, "<p class='heading1'>" & Listi(1).Nafn & "</p>"
  'Print #file, "<hr size=3 noshade color=#000000>"
  
  'Dim a() As Double
  'Dim b() As String
  'ReDim a(max) As Double
  'ReDim b(max) As String
  
  'For i = 1 To iTeljari
  '  If Listi(i).nr > 0 Then
  '    a(Listi(i).nr) = a(Listi(i).nr) + ((1 / 3) ^ (Listi(i).dypt + 2))
  '    b(Listi(i).nr) = Listi(i).Nafn
  '  End If
  'Next
  
  'For i = 1 To max
  '  If a(i) > 0 Then
  '    Print #file, b(i) & " " & Format(a(i), "0.000") & "%<br>"
  '  End If
  'Next
  
  Print #file, "<hr>"
  Print #file, "<table border=1><tr>"
  For i = 1 To iTeljari
    If Listi(i).fjoldi > 1 Then
      Print #file, "<td rowspan=" & Listi(i).fjoldi & " " & IIf(Listi(i).endurtekinn = True, "bgcolor='#ffffaf'", "") & "><p>"
      If Listi(i).mynd2 <> "" Then
        Print #file, "<img src=""" & pStillingar.Myndir & "/" & Listi(i).mynd2 & """ width='80' align='left'>"
      ElseIf Listi(i).mynd1 <> "" Then
        Print #file, "<img src=""" & pStillingar.Myndir & "/" & Listi(i).mynd1 & """ width='80' align='left'>"
      End If
      Print #file, "<a href=" & Listi(i).nr & ".htm>"
      Print #file, IIf(Listi(i).titill <> "", "<font color=red>" & Listi(i).titill, "")
      Print #file, Listi(i).Nafn & "</a><br>"
      Print #file, IIf(Listi(i).titill <> "", "</font>", "")
      Print #file, Listi(i).aettbokarnr & "<br>" & Listi(i).fdags & vbCrLf & "</p></td>"
    ElseIf Listi(i).nr > 0 Then
      Print #file, "<td " & IIf(Listi(i).endurtekinn = True, "bgcolor='#ffffaf'", "") & "><p>"
      If Listi(i).mynd2 <> "" Then
        Print #file, "<img src=""" & pStillingar.Myndir & "/" & Listi(i).mynd2 & """ width='80' align='left'>"
      ElseIf Listi(i).mynd1 <> "" Then
        Print #file, "<img src=""" & pStillingar.Myndir & "/" & Listi(i).mynd1 & """ width='80' align='left'>"
      End If
      Print #file, "<a href=" & Listi(i).nr & ".htm>"
      Print #file, IIf(Listi(i).titill <> "", "<font color=red>" & Listi(i).titill, "")
      Print #file, Listi(i).Nafn & "</a><br>"
      Print #file, IIf(Listi(i).titill <> "", "</font>", "")
      Print #file, Listi(i).aettbokarnr & "<br>" & Listi(i).fdags & vbCrLf
      Print #file, "</p></td><tr>" & vbCrLf
    Else
      Print #file, "<td><p>" '& "<a href=" & Listi(i).nr & ".htm>"
      'Print #file, IIf(Listi(i).titill <> "", "<font color=red>" & Listi(i).titill, "")
      'Print #file, Listi(i).Nafn & "&nbsp;" '& "</a><br>"
      'Print #file, IIf(Listi(i).titill <> "", "</font>", "")
      'Print #file, Listi(i).aettbokarnr & "<br>" & Listi(i).fdags & vbCrLf
      Print #file, "</p></td><tr>" & vbCrLf
    End If
  Next
  Print #file, "</table>"
  Print #file, "<hr><p>Höfundarréttur © 2000-" & Year(Now()) & " CaniGree.net<br>"
  Print #file, "<i>" & Now() & "</i></p>"
  Print #file, "</body></html>"
  
  Close #file
  Shell "explorer " & pStillingar.Mappa & "t" & numer & ".htm", vbNormalFocus
End Sub

Function forHundur(numer As Integer, dypi As Integer, maxdypi As Integer, Endur As Integer) As Long
  Dim rs As Recordset
  Dim i As Long
  Dim j As Long
  Dim nr, titill, aettbokarnr, fadir, modir, Nafn, fdags, mynd1, mynd2
  Dim endurtekinn As Boolean
  Dim iStadur As Long
  
  If dypi > maxdypi Then
    iTeljari = iTeljari + 1
    iStadur = iTeljari
    Listi(iStadur).nr = 0
    Listi(iStadur).Nafn = ""
    Listi(iStadur).titill = ""
    Listi(iStadur).aettbokarnr = ""
    Listi(iStadur).fdags = ""
    Listi(iStadur).fjoldi = 0
    Listi(iStadur).dypt = dypi
    Listi(iStadur).mynd1 = ""
    Listi(iStadur).mynd2 = ""
    forHundur = 1
    Exit Function
  End If
  
  Set rs = m_db.OpenRecordset("select * from hundalisti where nr=" & numer)
  nr = rs!nr
  fadir = "" & rs!fadirnr
  modir = "" & rs!modirnr
  Nafn = "" & rs!Nafn
  titill = "" & rs!titill
  fdags = "" & rs!fdags
  aettbokarnr = "" & rs!aettbokarnr
  mynd1 = "" & rs!mynd1
  mynd2 = "" & rs!mynd2
  rs.Close
  
  i = 0
  
  iTeljari = iTeljari + 1
  iStadur = iTeljari
  
  DoEvents
  
  If bBuinn(numer) = False Or Endur = 1 Then
    If bBuinn(numer) = True Then
    Listi(iStadur).endurtekinn = True
    For j = 0 To iStadur - 1
      If Listi(j).nr = nr Then
        Listi(j).endurtekinn = True
      End If
    Next
    End If
    
    bBuinn(numer) = True
    
    If Val(fadir) > 0 Then
      i = i + forHundur(Val(fadir), dypi + 1, maxdypi, Endur)
    ElseIf Val(modir) > 0 Then
      iTeljari = iTeljari + 1
      Listi(iTeljari).nr = 0
      Listi(iTeljari).Nafn = "&nbsp;"
      Listi(iTeljari).fjoldi = 1
      Listi(iTeljari).titill = ""
      Listi(iTeljari).aettbokarnr = ""
      Listi(iTeljari).fdags = ""
      Listi(iTeljari).dypt = dypi
      Listi(iTeljari).mynd1 = ""
      Listi(iTeljari).mynd2 = ""
      i = i + 1
    End If
    
    If Val(modir) > 0 Then
      i = i + forHundur(Val(modir), dypi + 1, maxdypi, Endur)
    ElseIf Val(fadir) > 0 Then
      iTeljari = iTeljari + 1
      Listi(iTeljari).nr = 0
      Listi(iTeljari).Nafn = "&nbsp;"
      Listi(iTeljari).fjoldi = 1
      Listi(iTeljari).titill = ""
      Listi(iTeljari).aettbokarnr = ""
      Listi(iTeljari).fdags = ""
      Listi(iTeljari).dypt = dypi
      Listi(iTeljari).mynd1 = ""
      Listi(iTeljari).mynd2 = ""
      i = i + 1
    End If
    
    Listi(iStadur).nr = nr
    Listi(iStadur).Nafn = Nafn
    Listi(iStadur).fjoldi = i
    Listi(iStadur).titill = titill
    Listi(iStadur).aettbokarnr = aettbokarnr
    Listi(iStadur).fdags = fdags
    Listi(iStadur).dypt = dypi
    Listi(iStadur).mynd1 = mynd1
    Listi(iStadur).mynd2 = mynd2
    
    If Val(fadir) = 0 And Val(modir) = 0 Then
      i = i + 1
    End If
  Else
    Listi(iStadur).endurtekinn = True
    For j = 0 To iStadur - 1
      If Listi(j).nr = nr Then
        Listi(j).endurtekinn = True
      End If
    Next
    
    Listi(iStadur).nr = nr
    Listi(iStadur).Nafn = Nafn & " (e)"
    Listi(iStadur).fjoldi = 0
    Listi(iStadur).titill = titill
    Listi(iStadur).aettbokarnr = aettbokarnr
    Listi(iStadur).fdags = fdags
    Listi(iStadur).dypt = dypi
    Listi(iStadur).mynd1 = mynd1
    Listi(iStadur).mynd2 = mynd2
    i = i + 1
  End If
  
  forHundur = i
End Function

'*******************************
'Almennar aðgerðir
'*******************************

Function listahund(rs As Recordset, span As Integer, skyld As String, Optional ffdags As Boolean = True, Optional breidd As Integer) As String
  Dim str As String
  Dim strAettbok As String
  On Error Resume Next
  If breidd < 1 Then breidd = 33
  If span > 1 Then
    str = "<td rowspan=" & span & " width=" & breidd & "%>"
  Else
    str = "<td width=" & breidd & "%>"
  End If
  
  Dim rs2 As Recordset
  Dim str2 As String
  Dim strFram As String
  Set rs2 = m_db.OpenRecordset("select * from hundar where nr=" & rs.Fields("nr" & skyld))
  If Not rs2.EOF Then
    strFram = rs2!latinn
    strAettbok = rs2!aettbokarnr
    If rs2!mynd2 & "" <> "" Then
      str = str & "<img src=""" & pStillingar.Myndir & "/" & rs2!mynd2 & """ width='80' align='left'>"
    ElseIf rs2!mynd1 & "" <> "" Then
      str = str & "<img src=""" & pStillingar.Myndir & "/" & rs2!mynd1 & """ width='80' align='left'>"
    End If
  End If
  rs2.Close
  
  str = str & "<p><a href=" & rs.Fields("nr" & skyld) & ".htm>"
On Error Resume Next
'  If sNoNulls(rs.Fields("my" & skyld)) <> "" Then
'    str = str & "<center><img src=" & pStillingar.Myndir & "/" & rs.Fields("my" & skyld) & "></center>"
'  End If
  If sNoNulls(rs.Fields("t" & skyld)) <> "" Then str = str & "<font color=red>" & rs.Fields("t" & skyld) & " "
  str = str & rs.Fields("n" & skyld)
  If sNoNulls(rs.Fields("t" & skyld)) <> "" Then str = str & "</font>"
  str = str & "</a>"
  If "" & strAettbok <> "" Then str = str & "<br>" & strAettbok
  If strFram = "fram" Then
    str = str & "<br><i>(Framræktun)</i></p>"
  ElseIf ffdags = True Then
    str = str & "<br>" & rs.Fields("fd" & skyld) & "</p>"
  Else
    str = str
  End If
On Error GoTo 0
  str = str & "</td>"
  If span = 1 Then str = str & "<tr>"
  listahund = str
End Function

Function listahund2(nr, titill, Nafn, fdags, span As Integer, breidd As Integer) As String
  Dim str As String
  Dim strAettbok As String
  On Error Resume Next
  If span > 1 Then
    str = "<td rowspan=" & span & " width=" & breidd & "%><p>"
  Else
    str = "<td width=" & breidd & "%><p>"
  End If
  
  Dim rs2 As Recordset
  Dim str2 As String
  Dim strFram As String
  Set rs2 = m_db.OpenRecordset("select * from hundar where nr=" & nr)
  If Not rs2.EOF Then
    strFram = rs2!latinn
    strAettbok = rs2!aettbokarnr
    str2 = rs2!mynd2 & ""
    If str2 <> "" Then
      str = str & "<img src=""" & pStillingar.Myndir & "/" & str2 & """ width='80'><br>"
    End If
  End If

  
  str = str & IIf(nr > 0, "<a href=" & nr & ".htm>", "")
  If "" & titill <> "" Then str = str & "<font color=red>" & titill & " "
  str = str & Nafn
  If "" & titill <> "" Then str = str & "</font>"
  str = str & IIf(nr > 0, "</a>", "")
  
  If "" & strAettbok <> "" Then str = str & "<br>" & strAettbok
  
  If strFram = "fram" Then
    str = str & "<br><i>(Framræktun)</i><br>"
  ElseIf "" & fdags <> "" Then
    str = str & "<br>" & fdags
  End If
  rs2.Close
  
  str = str & "</p></td>"
  If span = 1 Then str = str & "<tr>"
  listahund2 = str
End Function

Public Function reiknaIC(fnr As Long, Optional nrf As Long, Optional nrm As Long) As Double
  Dim rs, rs2 As Recordset
  Dim fic() As tIC
  Dim mic() As tIC
  Dim maxid As Long
  Dim i As Long
  Dim alls As Double
  Dim n As Integer, m As Integer
  
  If (dulkoda(pStillingar.Email, pStillingar.Leyfi) = False) Then
    MsgBox "Ekki er hægt að finna skyldleika í reynsluútgáfu forritsins.", vbOKOnly, "IC-gildi"
    reiknaIC = -1
    Exit Function
  End If
  
  Set rs = m_db.OpenRecordset("select max(id) as maxid from icgildi")
  maxid = rs!maxid + 1
  ReDim fic(maxid) As tIC
  ReDim mic(maxid) As tIC
  If Not bHundaIC Or m_countHundar < maxid Then
    ReDim m_hundaIC(maxid) As Double
    ReDim m_hundar(maxid) As tHundar
    m_countHundar = maxid
  
    Set rs2 = m_db.OpenRecordset("icgildi")
    While Not rs2.EOF
      m_hundar(rs2!id).sireID = rs2!sireID
      m_hundar(rs2!id).damID = rs2!damID
      rs2.MoveNext
    Wend
    rs2.Close
    bHundaIC = True
  End If
  rs.Close
  
  If fnr = 0 And nrf > 0 And nrm > 0 Then
    fnr = maxid
    m_hundar(fnr).sireID = nrf
    m_hundar(fnr).damID = nrm
  End If
  
  finnaIC m_hundar(fnr).sireID, fic, m_hundar, 1, 6, "F", fic, fnr
  finnaIC m_hundar(fnr).damID, mic, m_hundar, 1, 6, "M", fic, fnr
  
  For i = 1 To maxid
    If fic(i).fjoldi = 0 Or mic(i).fjoldi = 0 Then
      fic(i).fjoldi = 0
      mic(i).fjoldi = 0
    End If
  Next
  
  alls = 0
  
   
  For i = 1 To maxid
    If fic(i).fjoldi > 0 And mic(i).fjoldi > 0 Then
      For n = 1 To fic(i).fjoldi
        For m = 1 To mic(i).fjoldi
          If fic(i).afkvaemi(n) <> mic(i).afkvaemi(m) Then
            alls = alls + ((0.5 ^ (fic(i).lidur(n) + mic(i).lidur(m) - 1)) * (1 + IIf(m_hundaIC(i) > 0, m_hundaIC(i), reiknaIC(i))))
          End If
        Next
      Next
    End If
  Next
  
  m_hundaIC(fnr) = alls
  reiknaIC = alls
End Function

Private Sub finnaIC(fnr As Long, ByRef IC() As tIC, ByRef hundar() As tHundar, dypt As Long, maxdypt As Long, leggur As String, fic() As tIC, afkvaemi As Long)
   DoEvents
   If dypt > maxdypt Then Exit Sub

   If leggur = "M" And fic(fnr).fjoldi > 0 Then
     Dim i As Integer
     Dim b As Boolean
     b = True
     For i = 1 To fic(fnr).fjoldi
       If fic(fnr).afkvaemi(i) <> afkvaemi Then b = False
     Next
     If b Then Exit Sub
   End If
   
   IC(fnr).fjoldi = IC(fnr).fjoldi + 1
   IC(fnr).lidur(IC(fnr).fjoldi) = dypt
   IC(fnr).afkvaemi(IC(fnr).fjoldi) = afkvaemi
   
   If hundar(fnr).sireID > 0 Then
     finnaIC hundar(fnr).sireID, IC, hundar, dypt + 1, maxdypt, leggur, fic, fnr
   End If
   If hundar(fnr).damID > 0 Then
     finnaIC hundar(fnr).damID, IC, hundar, dypt + 1, maxdypt, leggur, fic, fnr
   End If
End Sub

Public Function forfedurIC(fnr As Long, Optional nrf As Long, Optional nrm As Long, Optional milli As String = "<br>") As String
  Dim rs, rs2 As Recordset
  Dim fic() As tIC
  Dim mic() As tIC
  Dim maxid As Long
  Dim i As Long
  Dim alls As String
  Dim allsIC As Double
  Dim n As Integer, m As Integer
  
  If (dulkoda(pStillingar.Email, pStillingar.Leyfi) = False) Then
    MsgBox "Ekki er hægt að finna skyldleika í reynsluútgáfu forritsins.", vbOKOnly, "IC-gildi"
    forfedurIC = ""
    Exit Function
  End If
  
  Set rs = m_db.OpenRecordset("select max(id) as maxid from icgildi")
  maxid = rs!maxid + 1
  ReDim fic(maxid) As tIC
  ReDim mic(maxid) As tIC
  If Not bHundaIC Or m_countHundar < maxid Then
    ReDim m_hundaIC(maxid) As Double
    ReDim m_hundar(maxid) As tHundar
    m_countHundar = maxid
  
    Set rs2 = m_db.OpenRecordset("icgildi")
    While Not rs2.EOF
      m_hundar(rs2!id).sireID = rs2!sireID
      m_hundar(rs2!id).damID = rs2!damID
      rs2.MoveNext
    Wend
    rs2.Close
    bHundaIC = True
  End If
  rs.Close
  
  If fnr = 0 And nrf > 0 And nrm > 0 Then
    fnr = maxid
    m_hundar(fnr).sireID = nrf
    m_hundar(fnr).damID = nrm
  End If
  
  finnaIC m_hundar(fnr).sireID, fic, m_hundar, 1, 6, "F", fic, fnr
  finnaIC m_hundar(fnr).damID, mic, m_hundar, 1, 6, "M", fic, fnr
  
  For i = 1 To maxid
    If fic(i).fjoldi = 0 Or mic(i).fjoldi = 0 Then
      fic(i).fjoldi = 0
      mic(i).fjoldi = 0
    End If
  Next
  
  alls = ""
   
  For i = 1 To maxid
    If fic(i).fjoldi > 0 And mic(i).fjoldi > 0 Then
      allsIC = 0
      For n = 1 To fic(i).fjoldi
        For m = 1 To mic(i).fjoldi
          If fic(i).afkvaemi(n) <> mic(i).afkvaemi(m) Then
            allsIC = allsIC + ((0.5 ^ (fic(i).lidur(n) + mic(i).lidur(m) - 1)) * (1 + IIf(m_hundaIC(i) > 0, m_hundaIC(i), reiknaIC(i))))
          End If
        Next
      Next
      Set rs = m_db.OpenRecordset("select nafn from hundalisti where nr=" & i)
      alls = alls & rs!Nafn & " - " & Format(allsIC, "0.0000%") & milli
    End If
  Next
  
  'alls = alls & "</p>"
  
  'm_hundaIC(fnr) = alls
  forfedurIC = alls
End Function


Public Function sNoNulls(str) As String
  sNoNulls = ""
  If Not IsNull(str) Then
    sNoNulls = str
  End If
End Function

