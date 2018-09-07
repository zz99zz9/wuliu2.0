<%
Exh_code=request.QueryString("Exh_code")
ECount=request.QueryString("ECount")
Exh_id=request.QueryString("Exh_id")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>支出费用表</title>
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
<%Response.AddHeader "content-disposition","attachment;filename="&Exh_code&"费用清算.xls"%>
  <table width="980" border="0" cellspacing="1" cellpadding="0" class="datalist">
    <tr>
      <th>项目编号</th>
      <th>展会名称</th>
      <th>项目主管</th>
      <th>应收</th>
      <th>已收</th>
      <th>应付</th>
      <th>已付</th>
      <th>预计利润</th>
      <th>目前利润</th>
      <th>未收款</th>
	</tr>
	<% Server.ScriptTimeOut=950 %>
    <%  
'开始分页


'打开数据库  
set rs=server.createobject("adodb.recordset")
sql="select * from Exhibition where Exh_id="&int(Exh_id)&" order by Exh_id desc"
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
do while not rs.eof 
%> 
    <tr onmousemove="changeTrColor(this)">
      <td><%=rs("Exh_Code")%></td>
      <td><%=rs("Exh_name")%></td>
      <td><%call Show_Supervisor_name(int(rs("Exh_Supid")))%></td>
      <td><%=Revenue_sum1(int(rs("Exh_id")))%></td>
<td><%=Revenue_sum2(int(rs("Exh_id")))%></td>
<td><%=Expense_sum1(int(rs("Exh_id")))%></td>
<td><%=Expense_sum2(int(rs("Exh_id")))%></td>
      <td><%=yjlr(int(rs("Exh_id")))%></td>
     <td><%=mqlr(int(rs("Exh_id")))%></td>
<td><%=wsk(int(rs("Exh_id")))%></td>
    </tr>
    <%
	 Revenue_sum111=Revenue_sum111+int(Revenue_sum11)
	 Revenue_sum222=Revenue_sum222+int(Revenue_sum22)
	 Expense_sum111=Expense_sum111+int(Expense_sum11)
	 Expense_sum222=Expense_sum222+int(Expense_sum22)
	 yjlrhz11=yjlrhz11+int(yjlrhz)
	 mqlr11=mqlr11+int(mqlrhz)
	 wsk11=wsk11+int(wskhz)
 rs.movenext


 count=count+1
 loop
 
end if
rs.close
set rs=nothing
%> 
<!--    <tr >
      <td></td>
      <td></td>
      <td></td>
      <td><%=Revenue_sum111%></td>
<td><%=Revenue_sum222%></td>
<td><%=Expense_sum111%></td>
<td><%=Expense_sum222%></td>
      <td><%=yjlrhz11%></td>
     <td><%=mqlr11%></td>
<td><%=wsk11%></td>
    </tr>-->
  </table>
</body>
</html>
