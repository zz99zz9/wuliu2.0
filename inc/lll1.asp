<html><head><title></title>
<meta http-equiv="refresh" >

</head>
<body >

<!--#include virtual="inc/conn.asp"-->
<script language="JavaScript" type="text/javascript"> 
function ChangeDiv(divId,divName,divName2,zDivCount) 
{ 
for(i=0;i<=zDivCount;i++) 
{ 
document.getElementById(divName+i).style.display="none"; 
document.getElementById(divName2+i).className="l"; 
//将所有的层都隐藏 
} 
document.getElementById(divName+divId).style.display="block"; 
document.getElementById(divName2+divId).className="l1"; 
//显示当前层 
} 
</script> 
<link rel="stylesheet" type="text/css" href="css/Public.css"/>
<style type="text/css">
body {
	background-color: #F0FAFB;
}
</style>
<%set rs=server.createobject("adodb.recordset")
sql="select * from Exhibition order by Exh_id desc"

rs.open sql,conn,3,3
count=0
do while not rs.eof %>

<div id="zhbb<%=count%>" onclick="javascript:ChangeDiv('<%=count%>','zhbm','zhbb',1)">
  <table width="100%" border="0" cellpadding="5" cellspacing="0">
    <tr><td width="20"></td><td width="19"><img src="../images/bao.gif" width="18" height="14" /></td>
  <td width="60"><%=rs("Exh_code")%></td>
  <td width="161"><!--<img src="images/f1.png" width="16" height="16" />--></td>
</tr>
</table></div>
  <div id="zhbm<%=count%>" class="zhbm" <%if count>0 then%> style="display:none;"<%end if%>><table width="100%" border="0" cellpadding="5" cellspacing="0">
    <tr>
      <td width="9%" height="25"></td>
      <td width="11%" rowspan="2" align="center" valign="top"><img src="../images/b3.gif" width="9" height="45" /></td>
      <td width="9%"><img src="../images/b1.gif" width="18" height="15" /></td>
      <td width="71%"><a href="#">业务结算</a></td>
      </tr>
    <tr>
      <td></td>
      <td><img src="../images/b2.gif" width="18" height="15" /></td>
      <td><a href="#">会展信息</a></td>
      </tr>
      
  </table></div>
  <%
 rs.movenext
 count=count+1
 loop
rs.close
set rs=nothing
%> </body></html>