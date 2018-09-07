<%
Exh_code=request.QueryString("Exh_code")
ECount=request.QueryString("ECount")
Exh_id=request.QueryString("Exh_id")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>应收款报表</title>
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
<%Response.AddHeader "content-disposition","attachment;filename="&Exh_code&"应收款报表.xls"%>
 <table width="980" border="0" cellspacing="1" cellpadding="0" class="datalist">
    <tr>
      <th>业务编号</th>
      <th>展会名称</th>
      <th>客户名称</th>
            <th>发票抬头</th>
      <th>费用项目</th>
      <th>应收</th>
      <th>已收</th>
      <th>欠款</th>

      <th>项目主管</th>
    </tr>
    <% Server.ScriptTimeOut=950 %>
   <%  
'开始分页


'打开数据库  
set rs=server.createobject("adodb.recordset")
sql="select * from Revenue order by Rev_id desc"
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
if rs("Rev_amount2")<>rs("Rev_amount1") then
%> 
    <tr onmousemove="changeTrColor(this)" <%if rs("Rev_amount1")<=rs("Rev_amount2") and rs("Rev_amount1")<>0 then%> style="display:hidden;"<%end if%>>

      <td><%call Show_exh_code(rs("Rev_exhid"))%></td>
      <td><%call Show_exh_name(rs("Rev_exhid"))%></td>
      <td><%call Show_customer_name(int(rs("Rev_customer")))%></td>
            <td><%=rs("Rev_Invoicename")%></td>
      <td><%call Show_Subject_name(int(rs("Rev_project")))%></td>
      <td><%=FormatNumber(rs("Rev_amount1"),2)%></td>
<td><%if rs("Rev_amount2")="" or rs("Rev_amount2")=0 then%>0.00<%else%><%=FormatNumber(rs("Rev_amount2"),2)%><%end if%></td>
<td><%=FormatNumber(rs("Rev_amount1")-rs("Rev_amount2"),2)%></td>


     
     <td><%call Show_exh2mas_name(rs("Rev_exhid"))%></td>
    </tr>
    <%		ys=ys+rs("Rev_amount1")
	yf=yf+rs("Rev_amount2")
		end if

	ye=ye+rs("Rev_amount1")-rs("Rev_amount2")
 rs.movenext
 count=count+1
 loop
end if
rs.close
set rs=nothing
if ys="" or ys=0 then
ys="0.00"
else
ys=FormatNumber(ys)
end if

if yf="" or yf=0 then
yf="0.00"
else
yf=FormatNumber(yf)
end if

if ye="" or ye=0 then
ye="0.00"
else
ye=FormatNumber(ye)
end if
%> 
    <tr>
         <td></td>
              <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td><%=ys%></td>
<td><%=yf%></td>
<td><%=ye%></td>

      <td></td>


    </tr>
  </table>
</body>
</html>
