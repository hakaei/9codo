<%
if InStr(Request.ServerVariables("HTTP_REFERER"), Request.ServerVariables("SERVER_NAME")) < 1 then
  response.write "^_^"
  response.end
end if
comments = "<!DOCTYPE HTML PUBLIC ""-//IETF//DTD HTML//EN"">"
comments = comments & "<html><head>"
comments = comments & "<meta http-equiv=""Content-Type"" content=""text/html; charset=big5"">"
comments = comments & "<title>行銷回應信函</title></head>"
comments = comments & "<body bgcolor=#ffffff>" & chr(13)
comments = comments & "<table border=0 cellspacing=0 cellpadding=0><tr><td bgcolor=#000000><table border=0 cellspacing=1 cellpadding=8>" & chr(13)
For Each item In Request.Form
  if item <> "submit" then
     if InStr(lcase(trim(Request.Form(item))), "[url") > 0 then
       response.write "&nbsp;"
       response.end
     elseif InStr(lcase(trim(Request.Form(item))), "http://") > 0 then
       response.write "請勿輸入網址"
       response.end
     end if
     if Request.Form(item) <> "0" then comments = comments & "<tr valign=top><td bgcolor=#ffc60f align=right><font color=#990000>" & Replace(item, "TotalPrice", "金額總計：") & "</td><td bgcolor=#ffffff>" & Replace(Request.Form(item),chr(13) & chr(10),"<br>") & "</td></tr>" & chr(13)
  end if
Next
comments = comments & "</table></td></tr></table>"
comments = comments & "</body></html>"

  ' 用 CDONTS
  Set smtp = Server.CreateObject("CDONTS.NewMail")
  smtp.BodyFormat = 0 
  smtp.MailFormat = 0 
  smtp.subject = "大象山食品網路訂購單"
  smtp.from = "snack@worldsnack.com.tw"
  smtp.To = "snack@worldsnack.com.tw"
  smtp.Body = comments
  On Error Resume Next
  smtp.Send
  if Err <> 0 then
    errstr = replace("信件傳送失敗: " & Err.Description, chr(13) & chr(10), "")
  else
    errstr = "訂單已收到, 謝謝您!"
    for i = 1 to 50
      session("Q_" & i) = 0
    next
  end if
  set smtp = nothing

response.write "<html><body bgcolor=#ffc60f><script Language=""JavaScript"">alert(""" & errstr & """);location.href='orderway.asp';</script></body></html>"
%>
