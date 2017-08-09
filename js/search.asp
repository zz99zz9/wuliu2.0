<!--#include file="../inc/conn.asp"-->
<% 
 Response.ContentType="text/xml"
 search=Trim(request("search"))

 set rs=server.CreateObject("adodb.recordset")
sql="select top 15 [Cus_name] from Customer where [Cus_name] like '"&search&"%' order by Cus_OrderId desc,Cus_id"

rs.open sql,conn,1,1

 str="<?xml version=""1.0"" encoding=""gb2312""?>"&vbnewline
  str=str&"<root>"&vbnewline
 If rs.eof Then  
 Else
  i=1
  Do While Not rs.eof
   str=str&"<message id="""&i&""">"&vbnewline  
   str=str&"  <text>"&trim(rs("Cus_name"))&"</text>"&vbnewline
   str=str&"</message>"&vbnewline
  i=i+1
  rs.movenext
  loop
  End If  
  str=str&"</root>"
  rs.close
  response.write str
%>