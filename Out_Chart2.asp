<%
Exh_code=request.QueryString("Exh_code")
ECount=request.QueryString("ECount")
Exh_id=request.QueryString("Exh_id")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>ҵ��һ����</title>
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
<%Response.AddHeader "content-disposition","attachment;filename="&request.cookies("S_year")&"��"&request.cookies("S_moon")&"�µ�"&request.cookies("E_year")&"��"&request.cookies("E_moon")&"��ҵ��һ����.xls"%>
<table width="980" border="0" cellspacing="1" cellpadding="0" class="datalist">
    <tr>
      <th height="30" colspan="9"><%=request.cookies("S_year")%>��<%=request.cookies("S_moon")%>��-<%=request.cookies("E_year")%>��<%=request.cookies("E_moon")%>�� ҵ��һ����</th>
    </tr>
    <tr>
      <th>��Ŀ���</th>
      <th>չ������</th>
            <th>���</th>
      <th>��ͷ</th>
      <th>��Ŀ����</th>
            <th>���</th>
      <th>����</th>
      <th>Ӧ��</th>
      <th>����</th>
      <th>�Ѹ�</th>
      <th>Ԥ������</th>
      <th>Ŀǰ����</th>
      <th>δ�տ�</th>
    </tr>
    <%  
'��ʼ��ҳ

sql1="where Exh_year>="&request.cookies("S_year")&" and Exh_moon>="&request.cookies("S_moon")&" and Exh_year<="&request.cookies("E_year")&" and Exh_moon<="&request.cookies("E_moon")&""
'�����ݿ�  
set rs=server.createobject("adodb.recordset")
sql="select * from Exhibition "&sql1&" order by Exh_id desc"
rs.PageSize = 100 '�����趨ÿҳ��ʾ�ļ�¼��
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
do while not rs.eof and count<rs.PageSize
%> 
    <tr onMouseMove="changeTrColor(this)">
      <td><%=rs("Exh_Code")%></td>
      <td><%=rs("Exh_name")%></td>
            <td><%call Show_class_name(rs("Exh_class"))%></td>
      <td><%=trim(rs("Exh_mark"))%></td>
      <td><%call Show_Supervisor_name(int(rs("Exh_Supid")))%></td>
            <td><%=rs("Exh_volume")%></td>
  
     <td><%=rs("Exh_kg")%></td>
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