<!--#include virtual="inc/conn.asp"-->

<%
dim wtime
set rs=server.createobject("adodb.recordset")
  sql="select * from Exhibition order by Exh_id desc"
   rs.open sql,conn,3,3
 do while not rs.eof 
    if rs("exh_moon")<10 then
    wtime=rs("exh_year")&".0"&rs("exh_moon")
    else
    wtime=rs("exh_year")&"."&rs("exh_moon")
    end if
    rs("w_time")=wtime
    response.write wtime&"<br>"
    i=i+1
     rs.update
 rs.movenext
 loop
	 rs.close
	 set rs=nothing

%>