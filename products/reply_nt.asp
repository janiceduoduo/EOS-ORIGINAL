<html>


<!-- Mirrored from www.hihosting.hinet.net/download/sendmail.htm by HTTrack Website Copier/3.x [XR&CO'2003], Wed, 30 Apr 2003 06:44:51 GMT -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>EOS</title>
</head>

<body>

<%
If Request("Send") <> Empty Then
 Set mail = Server.CreateObject("CDONTS.NewMail")
 mail.To = "service@eos.com.tw"
 mail.From = Request("Email")
 mail.Subject = Request("Subject")
 mail.Body = Request("Content")
 mail.Send
 Response.Write "Thank your incoming letter!<br>" & vbcrlf

 End If
%>
<script language=vbscript>
	MsgBox "Thank your incoming letter!" 
    location.href = "../products.html"  
 </script> 
<!---
<h2>使用iis的CDONTS.NewMail傳送email</h2>
<hr>
<FORM Action="reply.asp" Method=POST>
<TABLE Border=1>
<TR><TD>收件者:</TD>
<TD><INPUT TYPE=TEXT Name=To Size=40></TD></TR> 
<TR><TD>寄件者:</TD>
<TD><INPUT TYPE=TEXT Name=From Size=40></TD></TR> 
<TR><TD>主旨:</TD>
<TD><INPUT TYPE=TEXT Name=Subject Size=40></TD></TR>
<TR><TD>內文:</TD>
<TD><TEXTAREA Name=Body Rows=8 Cols=60></TEXTAREA></TD></TR> 

</TABLE>
<INPUT Type=Submit Value="送出" Name="Send"> 
</FORM>
--->

</body>


<!-- Mirrored from www.hihosting.hinet.net/download/sendmail.htm by HTTrack Website Copier/3.x [XR&CO'2003], Wed, 30 Apr 2003 06:44:51 GMT -->
</html>
