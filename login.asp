<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��������ϵͳ</title>
<link rel="stylesheet" type="text/css" href="css/Public.css"/>
<link rel="stylesheet" type="text/css" href="css/login.css"/>

</head>

<body>
<!--#include virtual="inc/conn.asp"-->
<!--#include virtual="inc/md5.asp"-->
<%
Check=request("act")
Ope_name=request("Ope_name")
Ope_password=md5(request("Ope_password"))
if Check="check" then
'��֤�û����Ƿ����
	sql="select Ope_name,Ope_password,Ope_visitor,Ope_id from operator where Ope_name='"&Ope_name&"' and Ope_password='"&Ope_password&"'"  ' ��ѯ���ݿ����Ƿ����ظ���¼
	Set rs=Server.CreateObject("ADODB.Recordset") 

	'set rs = conn.execute(sql)
	rs.open sql,conn,0,1
	If not(rs.Eof And rs.Bof) Then
	'cookiess
	response.cookies("wuliuuser")=Ope_name
	response.cookies("wuliuv")=rs("Ope_visitor")
	response.cookies("wuliuid")=rs("Ope_id")
	response.Redirect("index.asp")
	
	response.end()
	else 
	Response.Write ("<script language='javascript'>alert('��"&Ope_name&"�����û��������ڣ����������');history.back(-1);</script>") ' ���ؽ�������б���ת��
	response.end()
	end if
	rs.close
	set rs=nothing
end if%>
<div id="win">
<div id="loginpic" ></div>
<div id="loginfrom">
<form name="wuliuform" method="post" action="" onSubmit="return checkform();">
  <table width="100%" border="0" cellspacing="8" cellpadding="0">
    <tr>
      <td width="70">&nbsp;</td>
      <td>�û���
        <input name="Ope_name" type="text" id="Ope_name" size="15" /></td>
      <td>����
        <input name="Ope_password" type="password" id="Ope_password" size="15" /></td>
      <td width="130"><input type="submit" name="button" id="button" value="������¼" /><input name="act" type="hidden" id="act" value="check" /></td>
    </tr>
  </table>
  </form>
</div>
</div>
</body>
</html>
