<%
Exh_code=request.QueryString("Exh_code")
ECount=request.QueryString("ECount")
Exh_id=request.QueryString("Exh_id")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>֧�����ñ�</title>
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
<%Response.AddHeader "content-disposition","attachment;filename="&Exh_code&"֧�����ñ�.xls"%>
 <table width="980" border="0" cellspacing="1" cellpadding="0" class="datalist">
    <tr>
      <th>��Ŀ���</th>
      <th>չ������</th>
      <th>���</th>
            <th>��Ŀ����</th>
      <th>չ����</th>
      <th>���</th>
      <th>����</th>

    </tr>
    <% Server.ScriptTimeOut=950 %>
    <%  
'��ʼ��ҳ


'�����ݿ�  
set rs=server.createobject("adodb.recordset")
sql="select * from Exhibition order by Exh_id desc"
rs.PageSize = 10000 '�����趨ÿҳ��ʾ�ļ�¼��
rs.CursorLocation = 3

rs.open sql,conn,3,3
if err.number<>0 then
				response.write "���ݿ�����ʱ������"
				end if
				if rs.eof And rs.bof then
       				Response.Write "<p align='center' > �Բ���û�в�ѯ������Ҫ����Ϣ��</p>"
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
    <tr onmousemove="changeTrColor(this)" >

      <td><%=rs("Exh_code")%></td>
      <td><%=rs("Exh_name")%></td>
      <td><%call Show_class_name(rs("Exh_class"))%></td>
      <td><%call Show_Supervisor_name(rs("Exh_supid"))%></td>
      <td><%call Show_exh_count(rs("Exh_id"))%></td>
      <td><%=rs("Exh_volume")%></td>
<td><%=rs("Exh_kg")%></td>

    </tr>
    <%	

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
