<!--#include virtual="head.asp"-->
<!---->
<div id="bottom">
<div id="leftmenu">
<ul>
<li class="l" onclick="javascript:location.href='Company.asp'">���㵥λ</li>
<li class="l" onclick="javascript:location.href='Manager.asp'">��Ŀ����</li>
<li class="l" onclick="javascript:location.href='Subject.asp'">������Ŀ</li>
<li class="l1" onclick="javascript:location.href='Class.asp'">�ࡡ����</li>
<li class="l" onclick="javascript:location.href='Income.asp'">���뷽ʽ</li>
</ul>
</div>
<div id="rightContent">
<div id="rightContentadd2">
<%

act=request("act")
id=request("id")
Cla_name=left(trim(request("Cla_name")),10)
if act="add" then
'��֤�û����Ƿ����

	sql="select Cla_name from Class where Cla_name='"&Cla_name&"'"  ' ��ѯ���ݿ����Ƿ����ظ���¼
	
	set rs = conn.execute(sql)

	If not(rs.Eof And rs.Bof) Then
	 Response.Write ("<script language='javascript'>alert('��"&Cla_name&"���������Ѿ����ڣ����޸�����');history.back(-1);</script>") ' ���ؽ�������б���ת��
	response.end()
	end if

'���
 set news=server.CreateObject("adodb.recordset")
     sql="select * from Class "
	 news.open sql,conn,3,3	 
	 news.addnew	 
	 news("Cla_name")=Cla_name
news("Cla_time")=now()
	 news("Cla_OpeID")=request.cookies("wuliuid") '�û�id
     news.update
	 news.close
	 set news=nothing
response.write "<script language='javascript'>alert('�ύ�ɹ�');location.href('Class.asp');</script>"
end if
'�޸�
if act="mod" then
'��֤�û����Ƿ����
	sql="select Cla_name from Class where Cla_name='"&Cla_name&"'and Cla_id<>"&id  ' ��ѯ���ݿ����Ƿ����ظ���¼
	
	set rs = conn.execute(sql)
	
	If not(rs.Eof And rs.Bof) Then
	 Response.Write ("<script language='javascript'>alert('��"&Cla_name&"���������Ѿ����ڣ����޸�����');history.back(-1);</script>") ' ���ؽ�������б���ת��
	response.end()
	end if
'
 set news=server.CreateObject("adodb.recordset")
     sql="select * from Class where Cla_id="&id

	 news.open sql,conn,3,3	
	  
	' news.addnew	 
	 news("Cla_name")=Cla_name
news("Cla_time")=now()
	 news("Cla_OpeID")=request.cookies("wuliuid") '�û�id
     news.update
	 news.close
	 set news=nothing
response.write "<script language='javascript'>alert('�޸ĳɹ�');location.href('Class.asp');</script>"
end if

'ɾ������
if act="d" then

dsql="delete from Class where Cla_id="&id
response.write dsql
conn.execute dsql
response.redirect"Class.asp"
end if


'�������
if act="search" then

  if Cus_name<>"" then 
  sqlSe="where Cla_name like '%"&Cla_name&"%'"
  end if
end if
%>




<%if act="m" and id<>""  then%>
<%set modrs=server.createobject("adodb.recordset")
modsql="select * from Class where Cla_Id="&id

modrs.open modsql,conn,3,3%>
<form name="wuliuform" method="post" action="?id=<%=id%>" onSubmit="return checkform();">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="181" height="65">&nbsp;</td>
      <td width="244">������� <input name="Cla_name" type="text" id="Cla_name" size="20" value="<%=trim(modrs("Cla_name"))%>"/><%Old_Cla_name=trim(modrs("Cla_name"))%></td>
      <td width="245"><p></p></td>
      <td width="200"><input type="submit" name="button" id="button" value="�޸��������" /><input type="hidden" name="act" value="mod" /></td>
    </tr>
  </table>
</form>
<%else%>
<form name="wuliuform" method="post" action="" onSubmit="return checkform();">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50" height="65">&nbsp;</td>
      <td width="180">�������</td><td><p> <input name="Cla_name" type="text" id="Cla_name" size="25" /></p></td>
      <td width="122"><p></p></td>
      <td width="198"><input type="submit" name="button" id="button" value="�������" /><input type="hidden" name="act" value="add" /></td>
    </tr>
  </table>
</form>
<%end if%>
<form name="wuliuform" method="post" action="">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50" height="65">&nbsp;</td>
      <td width="180">������� </td>
      <td><p>
        <input name="Cla_name" type="text" id="Cla_name" size="25" value="<%=Cus_name%>"/>
      </p>        <!--ȫ��
        <input name="Cus_name2" type="text" id="Cus_name2" size="20" />--></td>
      <td width="122"><p></p></td>
      <td width="198"><input type="submit" name="button" id="button" value="�������" /><input type="hidden" name="act" value="search" /></td>
    </tr>
  </table>
</form>
<script>
//�ύ��֤
function checkform()
{
     var reg = /[^\w\u4e00-\u9fa5]/g;    // \w�������֡���ĸ�����ִ�Сд�����»��ߡ���\u4e00-\u9fa5�����֡� 
  var oName = document.wuliuform.Cla_name.value;;
      if (oName==""){
	
	wuliuform.Cla_name.focus();
	
      return false;

    }
	else if(reg.test(oName)){
			wuliuform.Cla_name.focus();
	
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
  var aP = document.getElementsByTagName('p');
  var name_msg = aP[0];
   var name_length = 0;
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

  }


</script>
</div>
<div id="rightContentlist">
<table style="width:100%">
<tr><th>ϵͳID</th>
<th>���</th>
<th>���ʱ��</th><th>����</th></tr>
<%  
'��ʼ��ҳ

dim intPage,page,pre,last,filepath 
'�����ݿ�  
set rs=server.createobject("adodb.recordset")
sql="select * from Class "&sqlSe&" order by Cla_OrderId desc,Cla_id"
rs.PageSize = 50000 '�����趨ÿҳ��ʾ�ļ�¼��
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
<tr onmousemove="changeTrColor(this)"><td><%=rs("Cla_id")%></td><td><%=rs("Cla_name")%></td><td><%=formatdatetime(rs("Cla_time"),2)%></td><td><a href="?act=m&id=<%=rs("Cla_id")%>"><img src="images/m.gif" border=0/></a>��<a href="javascript:del()"><img src="images/d.gif" border=0/></a><SCRIPT LANGUAGE="JavaScript">
 <!-- 

 function del(){                   
 if(window.confirm("ȷʵҪɾ����<%=trim(rs("Cla_name"))%>����")){            
  window.location = "Class.asp?act=d&id=<%=rs("Cla_id")%>"; 
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
