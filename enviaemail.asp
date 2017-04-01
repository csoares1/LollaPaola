<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
response.Charset = "utf-8"
response.ContentType = "text/html"
Set Mail = CreateObject("CDO.Message")

Dim lngBytesCount
lngBytesCount = Request.TotalBytes
jsonstring = BytesToStr(Request.BinaryRead(lngBytesCount))
a=Split(jsonstring,",")

for each x in a
    mensagem = mensagem + x & "<br />"
    mensagem = replace(mensagem,"{","")
    mensagem = replace(mensagem,"}","")
    mensagem = replace(mensagem,Chr(34),"")
next

Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") ="smtp.gmail.com"
Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465

Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1

Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") ="orcamentolollapaoladecora@gmail.com"
Mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") ="Loll@0404"

Mail.Configuration.Fields.Update
Mail.BodyPart.Charset = "utf-8"

Mail.Subject="Orçamento LollaPaola"
Mail.From="orcamentolollapaoladecora@gmail.com"
Mail.To="orcamentolollapaoladecora@gmail.com"
Mail.HTMLBody ="</html>"&_
 "<head></head>"&_
 "<title></title>"&_
    "<body>"&_
          "</h1><p style='font-size:18px;'>"&_
          "Prezado(s)</p><br/>"&_
          "<p style='font-size:16px;'>Segue detalhes da solicitação de orçamento.</p>"&_
          "<table width='100%' cellspading='5'>"&_
            "<thead >"&_
              "<tr bgcolor='brown' style='font-size: 11px;font-family: Tahoma;color:white;'>"&_
              "<th><p style='font-size:14px;'  width='35%'>Mensagem</p></th>"&_
              "</tr>"&_
            "</thead>"&_
            "<tbody>"&_
              "<tr bgcolor='#E3EFFA' style='font-size: 11px;font-family: Tahoma;color:#3B3C3D;'>"&_
              "<td>"&_
              "<p style='font-size:16px;'>"&mensagem&"</p>"&_
              "</td>"&_
            "</tbody>"&_
          "</table>"&_
    "</body>"&_
"</html>"
Mail.Send
Set Mail = Nothing

Function BytesToStr(bytes)
      Dim Stream
      Set Stream = Server.CreateObject("Adodb.Stream")
          Stream.Type = 1 'adTypeBinary
          Stream.Open
          Stream.Write bytes
          Stream.Position = 0
          Stream.Type = 2 'adTypeText
          Stream.Charset = "iso-8859-1"
          BytesToStr = Stream.ReadText
          Stream.Close
      Set Stream = Nothing
  End Function



%>
