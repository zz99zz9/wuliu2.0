<%
Exh_code=request.QueryString("Exh_code")
ECount=request.QueryString("ECount")
Exh_id=request.QueryString("Exh_id")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>业务一览表</title>
<!--#include virtual="inc/conn.asp"-->
<!--#include virtual="inc/inc.asp"-->
<style type="text/css">
table {
	font-family: verdana,arial,sans-serif;
	font-size:11px;
	color:#333333;
	border-width: 1px;
	border-color: #666666;
	border-collapse: collapse;
}
table th {
	border-width: 1px;
	padding: 8px;
	border-style: solid;
	border-color: #666666;
	background-color: #dedede;
	height:35px;
	font-size:12px;
}
table td {
	border-width: 1px;
	padding: 8px;
	border-style: solid;
	border-color: #666666;
	background-color: #ffffff;
	height:35px;
	font-size:12px;
}
</style>
</head>

<body>
<%response.ContentType ="application/vnd.ms-excel"%> 
<%Response.AddHeader "content-disposition","attachment;filename="&request.cookies("S_year")&"年"&request.cookies("S_moon")&"月到"&request.cookies("E_year")&"年"&request.cookies("E_moon")&"月业务一览表.xls"%>
<table width="980" border="0" cellspacing="1" cellpadding="0" class="datalist">
    <tr>
      <th height="30" colspan="9"><%=request.cookies("S_year")%>年<%=request.cookies("S_moon")%>月-<%=request.cookies("E_year")%>年<%=request.cookies("E_moon")%>月 业务一览表</th>
    </tr>
    <tr>
	  <th>ID</th>
      <th>项目编号</th>
      <th>展会名称</th>
      <th>项目主管</th>
      <th>应收</th>
      <th>已收</th>
      <th>已付</th>
      <th>预计利润</th>
      <th>目前利润</th>
      <th>未收款</th>
	</tr>
	<% Server.ScriptTimeOut=950 %>
    <%  
'开始分页

s_time=FormatNumber(request.cookies("S_year")&"."&request.cookies("S_moon"),2,False,False,False)
e_time=FormatNumber(request.cookies("E_year")&"."&request.cookies("E_moon"),2,False,False,False)
sql1="where w_time>="&s_time&" and w_time<="&e_time&""
if request.cookies("Sup_id")<>"0" then
  sql1=sql1+" and Exh_Supid="&request.cookies("Sup_id")
end if
'打开数据库  
set rs=server.createobject("adodb.recordset")
sql="select * from Exhibition "&sql1&" order by Exh_id desc"
rs.PageSize = 10000 '这里设定每页显示的记录数
rs.CursorLocation = 3

rs.open sql,conn,3,3
if err.number<>0 then
				response.write "数据库中暂时无数据"
				end if
				if rs.eof And rs.bof then
       				Response.Write "<p align='center' > 对不起，没有查询到您需要的信息！</p>"
   				else
	  				pre = true
last = true
page = trim(Request.QueryString("page"))

if len(page) = 0 then
intpage = 1
pre = false
else
if cint(page) =< 1 then
intpage = 1
pre = false
else
if cint(page) >= rs.PageCount then
intpage = rs.PageCount
last = false
else
intpage = cint(page)
end if
end if
end if
if not rs.eof then
rs.AbsolutePage = intpage
end if 
i=0
do while not rs.eof and count<rs.PageSize
i=i+1
%> 
    <tr onMouseMove="changeTrColor(this)">
	<td><%=i%></td>
      <td><%=rs("Exh_Code")%></td>
      <td><%=rs("Exh_name")%></td>
      <td><%call Show_Supervisor_name(int(rs("Exh_Supid")))%></td>
      <td><%call Revenue_sum1(int(rs("Exh_id")))%></td>
<td><%call Revenue_sum2(int(rs("Exh_id")))%></td>
<td><%call Expense_sum1(int(rs("Exh_id")))%></td>

      <td><%call yjlr(int(rs("Exh_id")))%></td>
     <td><%call mqlr(int(rs("Exh_id")))%></td>
<td><%call wsk(int(rs("Exh_id")))%></td>
    </tr>
    <%
 rs.movenext
 count=count+1
 loop
end if
rs.close
set rs=nothing
%> 
  </table>
</body>
</html>
