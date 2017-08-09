<%
'统计数据条数
set rs=server.createobject("adodb.recordset")
sql="select count(Exh_id) as Count1 from Exhibition "

rs.open sql,conn,3,3
if not rs.eof then
count1=rs("Count1")
end if
rs.close
set rs=nothing


%>
<%
'读取数据
set rs=server.createobject("adodb.recordset")
sql="select top 20 * from Exhibition order by Exh_id desc"

rs.open sql,conn,3,3
count=0
do while not rs.eof %>

<div id="zhbbb<%=count%>" onClick="javascript:ChangeDiv('<%=count%>','zhbmm','zhbbb',<%=count1-1%>)">
  <table width="100%" border="0" cellpadding="5" cellspacing="0">
    <tr><td width="20"></td><td width="19"><img src="../images/bao.gif" width="18" height="14" /></td>
  <td width="60" class="zhbb"><a href="#"><%=rs("Exh_code")%></a></td>
  <td width="161"><!--<img src="images/f1.png" width="16" height="16" />--></td>
</tr>
</table></div>
  <div id="zhbmm<%=count%>" class="zhbm" <%if count<>int(ECount2) then%> style="display:none;"<%end if%>><table width="100%" border="0" cellpadding="5" cellspacing="0">
    <tr>
      <td width="9%" height="25"></td>
      <td width="11%" rowspan="2" align="center" valign="top"><img src="../images/b3.gif" width="9" height="45" /></td>
      <td width="9%"><img src="../images/b1.gif" width="18" height="15" /></td>
      <td width="71%"><a href="?Exh_code=<%=trim(rs("Exh_code"))%>&ECount2=<%=count%>&Exh_id=<%=trim(rs("Exh_id"))%>&zhlb=1">业务结算</a></td>
      </tr>
       <%if request.cookies("wuliuv")=0 then%>
    <tr>
      <td></td>
      <td><img src="../images/b2.gif" width="18" height="15" /></td>
      <td><a href="Edit_Exhibition.asp?Exh_id=<%=trim(rs("Exh_id"))%>">会展信息</a></td>
      </tr>
      <%end if%>
  </table></div>
  <%
 rs.movenext
 count=count+1
 loop
rs.close
set rs=nothing
%>
  