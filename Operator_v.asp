<!--#include virtual="head.asp"-->
<!---->
<div id="bottom">
<div id="leftmenu">
<ul>
<li class="l" onclick="javascript:location.href='Operator.asp'">����Ա����</li>
<li class="l1" onclick="javascript:location.href='Operator_v.asp'">����Ȩ������</li>

</ul>
</div>
<div id="rightContent">
<div id="rightContentadd">
<%

act=request("act")
id=request("id")
Ope_name=request("Ope_name")
Ope_password=md5(request("Ope_password"))
if act="add" then
'��֤�û����Ƿ����

	sql="select Ope_name from operator where Ope_name='"&Ope_name&"'"  ' ��ѯ���ݿ����Ƿ����ظ���¼
	
	set rs = conn.execute(sql)

	If not(rs.Eof And rs.Bof) Then
	 Response.Write ("<script language='javascript'>alert('��"&Ope_name&"�����û����Ѿ����ڣ����޸�����');history.back(-1);</script>") ' ���ؽ�������б���ת��
	response.end()
	end if

'���
 set news=server.CreateObject("adodb.recordset")
     sql="select * from operator "
	 news.open sql,conn,3,3	 
	 news.addnew	 
	 news("Ope_name")=Ope_name
	 news("Ope_password")=Ope_password
news("Ope_time")=now()
news("Ope_visitor")=1
'news("Sup_OpeID")=Sup_OpeID '�û�id
     news.update
	 news.close
	 set news=nothing
response.write "<script language='javascript'>alert('�ύ�ɹ�');location.href('operator_v.asp');</script>"
end if
'�޸�
if act="mod" then
'��֤�û����Ƿ����
	sql="select Ope_name from operator where Ope_name='"&Ope_name&"'and Ope_id<>"&id  ' ��ѯ���ݿ����Ƿ����ظ���¼
	
	set rs = conn.execute(sql)
	
	If not(rs.Eof And rs.Bof) Then
	 Response.Write ("<script language='javascript'>alert('��"&Ope_name&"�����û����Ѿ����ڣ����޸�����');history.back(-1);</script>") ' ���ؽ�������б���ת��
	response.end()
	end if
'
 set news=server.CreateObject("adodb.recordset")
     sql="select * from operator where Ope_id="&id

	 news.open sql,conn,3,3	
	  
	' news.addnew	 
	 news("Ope_name")=Ope_name
	 	 news("Ope_password")=Ope_password
		 news("Ope_visitor")=1
news("Ope_time")=now()
'news("Sup_OpeID")=Sup_OpeID '�û�id
     news.update
	 news.close
	 set news=nothing
response.write "<script language='javascript'>alert('�޸ĳɹ�');location.href('operator_v.asp');</script>"
end if

'ɾ������
if act="d" then

dsql="delete from operator where Ope_id="&id
response.write dsql
conn.execute dsql
response.redirect"operator.asp"
end if%>




<%if act="m" and id<>""  then%>
<%set modrs=server.createobject("adodb.recordset")
modsql="select * from operator where Ope_Id="&id

modrs.open modsql,conn,3,3%>
<form name="wuliuform" method="post" action="?id=<%=id%>" onSubmit="return checkform();">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
    <td width="53" height="65">&nbsp;</td>
      <td width="64">��¼����</td>
      <td width="140"><input name="Ope_name" type="text" id="Ope_name" size="20" value="<%=trim(modrs("Ope_name"))%>"/><%Old_sup_name=trim(modrs("Ope_name"))%></td>
      <td width="120"><p></p></td>
      <td width="37">����
        </td>
      <td width="140"><input name="Ope_password" type="password" id="Ope_password" size="20" value="<%=trim(modrs("Ope_password"))%>"/></td>
      <td width="133"><p></p></td>
      <td width="183"><input type="submit" name="button" id="button" value="�޸Ĳ���Ա" /><input type="hidden" name="act" value="mod" /></td>

    </tr>
  </table>
</form>
<%else%>
<form name="wuliuform" method="post" action="" onSubmit="return checkform();">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="53" height="65">&nbsp;</td>
      <td width="64">��¼����</td>
      <td width="140"><input name="Ope_name" type="text" id="Ope_name" size="20" /></td>
      <td width="120"><p></p></td>
      <td width="37">����
        </td>
      <td width="140"><input name="Ope_password" type="password" id="Ope_password" size="20" /></td>
      <td width="133"><p></p></td>
      <td width="183"><input type="submit" name="button" id="button" value="��������Ա" /><input type="hidden" name="act" value="add" /></td>
    </tr>

  </table>
</form>
<%end if%>
<script>
//�ύ��֤
function checkform()
{
     var reg = /[^\w\u4e00-\u9fa5]/g;    // \w�������֡���ĸ�����ִ�Сд�����»��ߡ���\u4e00-\u9fa5�����֡� 
  var oName = document.wuliuform.Ope_name.value;
  var Password = document.wuliuform.Ope_password.value;
      if (oName==""){
	
	wuliuform.Ope_name.focus();
	
      return false;

    }
	else if(reg.test(oName)){
			wuliuform.Ope_name.focus();
	
      return false;
		}
	else if (Password==""){
	
	wuliuform.Ope_password.focus();
	
      return false;

    }

    else{
	   return true;

    }
	
}

//������֤
window.onload=function(){
  var aInput = document.getElementsByTagName('input');
  var oName = aInput[0];
  var pwd = aInput[1];
	  
  var aP = document.getElementsByTagName('p');
  var name_msg = aP[0];
  var pwd_msg = aP[1];
 //  var name_length = 0;
//��Ա��

  oName.onfocus = function(){
    name_msg.style.display = "inline";
    name_msg.innerHTML = "<i class='info'></i>�Ƽ�ʹ��������";
  }

  oName.onblur = function(){

    //���зǷ��ַ�            
    var reg = /[^\w\u4e00-\u9fa5]/g;    // \w�������֡���ĸ�����ִ�Сд�����»��ߡ���\u4e00-\u9fa5�����֡� 


    //����Ϊ��
    if (this.value==""){
	
      name_msg.innerHTML = "<i class='no'></i>����Ϊ�գ�";
    }
    else if(reg.test(this.value)){
      name_msg.innerHTML = '<i class="no"></i>���зǷ��ַ���';
    }

    //OK
    else {
      name_msg.innerHTML = "<i class='yes'></i>OK��";

    }
  }
//����
  pwd.onfocus = function(){
    pwd_msg.style.display = "inline";
    pwd_msg.innerHTML = "<i class='info'></i>����6��12λ����";
  }

  pwd.onblur = function(){


    if (this.value==""){
	
      pwd_msg.innerHTML = "<i class='no'></i>����Ϊ�գ�";
    }

    //OK
    else {
      pwd_msg.innerHTML = "<i class='yes'></i>OK��";

    }


  }
  }


</script>
</div>
<div id="rightContentlist">
<table style="width:100%">
<tr><th>ϵͳID</th>
<th>����Ա</th>
<th>���ʱ��</th><th>����</th></tr>
<%  
'��ʼ��ҳ

dim intPage,page,pre,last,filepath 
'�����ݿ�  
set rs=server.createobject("adodb.recordset")
sql="select * from operator where Ope_visitor=1 order by Ope_id"
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
<tr onmousemove="changeTrColor(this)"><td><%=rs("Ope_id")%></td><td><%=rs("Ope_name")%></td><td><%=formatdatetime(rs("Ope_time"),2)%></td><td><a href="?act=m&id=<%=rs("Ope_id")%>"><img src="images/m.gif" border=0/></a>��<a href="javascript:del()"><img src="images/d.gif" border=0/></a><SCRIPT LANGUAGE="JavaScript">
 <!-- 

 function del(){                   
 if(window.confirm("ȷʵҪɾ����<%=trim(rs("Ope_name"))%>����")){            
  window.location = "operator.asp?act=d&id=<%=rs("Ope_id")%>"; 
  //�ύ��url         
  }else{             
  return;         
  }   
    } //--> 
    </SCRIPT></td></tr>
<%
 rs.movenext
 count=count+1
 loop
end if
%>  
</table>
<script type="text/javascript">
function changeTrColor(obj){    
    var _table=obj.parentNode;
    for (var i=0;i<_table.rows.length;i++){
        _table.rows[i].style.backgroundColor="";
    }    
    obj.style.backgroundColor="#E2CFBC";
}
</script>
</div>
</div>
</div>

</body>
</html>
