<!--#include virtual="inc/conn.asp"-->

<%
dim supid
set rs=server.createobject("adodb.recordset")
'del top 500
  sql="select * from Expense order by Exp_id desc"
   rs.open sql,conn,3,3
 do while not rs.eof 
    '读取年月写入wtime
    if isNumeric(rs("Exp_Exhid")) then
   '     set srs=server.createobject("adodb.recordset")
   '     ssql="select * from Exhibition where Exh_id="&rs("Exp_Exhid")
            response.write ssql&"<br>"
    '     srs.open ssql,conn,3,3
    '     supid=srs("exh_supid")

    'rs("supid")=supid
   response.write supid&"<br>"
    'i=i+1
    ' rs.update
    ' end if
 rs.movenext
 loop
	 rs.close
	 set rs=nothing

%>