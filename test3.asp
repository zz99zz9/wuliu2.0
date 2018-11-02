<!--#include virtual="inc/conn.asp"-->

<%
dim wtime
set rs=server.createobject("adodb.recordset")
'del top 500
  sql="select * from Revenue order by Rev_id desc"
   rs.open sql,conn,3,3
 do while not rs.eof 
    '读取年月写入wtime
    if isNumeric(rs("Rev_Exhid")) then
        set srs=server.createobject("adodb.recordset")
        ssql="select * from Exhibition where Exh_id="&rs("Rev_Exhid")
        response.write ssql&"<br>"
        srs.open ssql,conn,3,3
        wtime=srs("w_time")

    rs("w_time")=wtime
   ' response.write wtime&"<br>"
    i=i+1
     rs.update
     end if
 rs.movenext
 loop
	 rs.close
	 set rs=nothing

%>