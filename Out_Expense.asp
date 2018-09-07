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
<%Response.AddHeader "content-disposition","attachment;filename="&Exh_code&"支出费用表.xls"%>
 <table width="980" border="0" cellspacing="1" cellpadding="0" class="datalist">
    <tr>
      <th>编号</th>
      <th>客户名称</th>
      <th>费用项目</th>
      <th>应付</th>
<th>已付</th>
      <th>发票号</th>
      <th>发票抬头</th>
      <th>支出方式</th>
      <th>备注</th>
      <th>更改人</th>
      <th>更改日期</th>

    </tr>
    <% Server.ScriptTimeOut=950 %>
    <%  
'开始分页


'打开数据库  
set rs=server.createobject("adodb.recordset")
sql="select * from Expense where Exp_Exhid="&int(Exh_id)&" order by Exp_id desc"


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
    <tr onmousemove="changeTrColor(this)" <%if rs("Exp_amount1")<=rs("Exp_amount2") and rs("Exp_amount1")<>0 then%> style="color:#ff0000;"<%end if%>>
      <td><%=Exh_code%></td>
      <td><%call Show_customer_name(int(rs("Exp_customer")))%></td>
      <td><%call Show_Subject_name(int(rs("Exp_project")))%></td>
      <td><%=FormatNumber(rs("Exp_amount1"))%></td>
<td><%=FormatNumber(rs("Exp_amount2"))%></td>
      <td><%=rs("Exp_Invoiceid")%></td>
      <td><%=rs("Exp_Invoicename")%></td>
      <td><%call Show_Income_name(int(rs("Exp_mode")))%></td>
      <td><%=rs("Exp_content")%></td>
      <td><%call Show_operator_name(int(rs("Exp_Opeid")))%></td>
      <td><%=formatdatetime(rs("Exp_time"),2)%></td>

    </tr>
    <%
 rs.movenext
 count=count+1
 loop
end if
rs.close
set rs=nothing
%> 
<%if Exh_id<>"" then%>
<tr>
 <td>汇总</td>
      <td></td>
      <td></td>
      <td><b><%=Expense_sum1(Exh_id)%></b></td>
<td><b><%Expense_sum2(Exh_id)%></b></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>

</tr>
<%end if%>
  </table>
</body>
</html>
