<!--#include virtual="head.asp"-->
<!---->
<div id="bottom">
<div id="leftmenu">
<ul>
<li class="l" onclick="javascript:location.href='Company.asp'">���㵥λ</li>
<li class="l" onclick="javascript:location.href='Manager.asp'">��Ŀ����</li>
<li class="l1" onclick="javascript:location.href='Subject.asp'">������Ŀ</li>
<li class="l" onclick="javascript:location.href='Class.asp'">�ࡡ����</li>
<li class="l" onclick="javascript:location.href='Income.asp'">���뷽ʽ</li>
</ul>
</div>
<div id="rightContent">
<div id="rightContentadd2">
<%

act=request("act")
id=request("id")
Sub_name=request("Sub_name")
if act="add" then
'��֤�û����Ƿ����

	sql="select Sub_name from Subject where Sub_name='"&Sub_name&"'"  ' ��ѯ���ݿ����Ƿ����ظ���¼
	
	set rs = conn.execute(sql)

	If not(rs.Eof And rs.Bof) Then
	 Response.Write ("<script language='javascript'>alert('��"&Sub_name&"���������Ѿ����ڣ����޸�����');history.back(-1);</script>") ' ���ؽ�������б���ת��
	response.end()
	end if

'���
 set news=server.CreateObject("adodb.recordset")
     sql="select * from Subject "
	 news.open sql,conn,3,3	 
	 news.addnew	 
	 news("Sub_name")=Sub_name
news("Sub_time")=now()
	 news("Sub_OpeID")=request.cookies("wuliuid") '�û�id
     news.update
	 news.close
	 set news=nothing
response.write "<script language='javascript'>alert('�ύ�ɹ�');location.href('Subject.asp');</script>"
end if
'�޸�
if act="mod" then
'��֤�û����Ƿ����
	sql="select Sub_name from Subject where Sub_name='"&Sub_name&"'and Sub_id<>"&id  ' ��ѯ���ݿ����Ƿ����ظ���¼
	
	set rs = conn.execute(sql)
	
	If not(rs.Eof And rs.Bof) Then
	 Response.Write ("<script language='javascript'>alert('��"&Sub_name&"���������Ѿ����ڣ����޸�����');history.back(-1);</script>") ' ���ؽ�������б���ת��
	response.end()
	end if
'
 set news=server.CreateObject("adodb.recordset")
     sql="select * from Subject where Sub_id="&id

	 news.open sql,conn,3,3	
	  
	' news.addnew	 
	 news("Sub_name")=Sub_name
news("Sub_time")=now()
	 news("Sub_OpeID")=request.cookies("wuliuid") '�û�id
     news.update
	 news.close
	 set news=nothing
response.write "<script language='javascript'>alert('�޸ĳɹ�');location.href('Subject.asp');</script>"
end if

'ɾ������
if act="d" then

dsql="delete from Subject where Sub_id="&id
response.write dsql
conn.execute dsql
response.redirect"Subject.asp"
end if

'������������
if act="search" then

  if Sub_name<>"" then 
  sqlSe="where Sub_name like '%"&Sub_name&"%'"
  end if
end if
%>




<%if act="m" and id<>""  then%>
<%set modrs=server.createobject("adodb.recordset")
modsql="select * from Subject where Sub_Id="&id

modrs.open modsql,conn,3,3%>
<form name="wuliuform" method="post" action="?id=<%=id%>" onSubmit="return checkform();">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="181" height="65">&nbsp;</td>
      <td width="244">��������
        <input name="Sub_name" type="text" id="Sub_name" size="20" value="<%=trim(modrs("Sub_name"))%>"/><%Old_sub_name=trim(modrs("Sub_name"))%></td>
      <td width="245"><p></p></td>
      <td width="200"><input type="submit" name="button" id="button" value="�޸ķ�������" /><input type="hidden" name="act" value="mod" /></td>
    </tr>
  </table>
</form>
<%else%>
<form name="wuliuform" method="post" action="" onSubmit="return checkform();">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50" height="65">&nbsp;</td>
      <td width="180">������Ŀ����<td><td><p><input name="Sub_name" type="text" id="Sub_name" size="25" /></p></td>
      <td width="122"><p></p></td>
      <td width="198"><input type="submit" name="button" id="button" value="������������" /><input type="hidden" name="act" value="add" /></td>
    </tr>
  </table>
</form>
<%end if%>
<form name="wuliuform" method="post" action="">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50" height="65">&nbsp;</td>
      <td width="180">������Ŀ����</td>
      <td><p>
        <input name="Sub_name" type="text" id="Sub_name" size="25" value="<%=Cus_name%>"/>
      </p>        <!--ȫ��
        <input name="Cus_name2" type="text" id="Cus_name2" size="20" />--></td>
      <td width="122"><p></p></td>
      <td width="198"><input type="submit" name="button" id="button" value="������Ŀ����" /><input type="hidden" name="act" value="search" /></td>
    </tr>
  </table>
</form>
<script>
//�ύ��֤
function checkform()
{
     var reg = /[^\w\u4e00-\u9fa5]/g;    // \w�������֡���ĸ�����ִ�Сд�����»��ߡ���\u4e00-\u9fa5�����֡� 
  var oName = document.wuliuform.Sub_name.value;;
      if (oName==""){
	
	wuliuform.Sub_name.focus();
	
      return false;

    }
	else if(reg.test(oName)){
			wuliuform.Sub_name.focus();
	
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
<th>��������</th>
<th>���ʱ��</th><th>����</th></tr>
<%  
'��ʼ��ҳ

dim intPage,page,pre,last,filepath 
'�����ݿ�  
set rs=server.createobject("adodb.recordset")
sql="select * from Subject "&sqlSe&" order by Sub_OrderId desc,Sub_id"
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
<tr onmousemove="changeTrColor(this)"><td><%=rs("Sub_id")%></td><td><%=rs("Sub_name")%></td><td><%=formatdatetime(rs("Sub_time"),2)%></td><td><a href="?act=m&id=<%=rs("Sub_id")%>"><img src="images/m.gif" border=0/></a>��<a href="javascript:del()"><img src="images/d.gif" border=0/></a><SCRIPT LANGUAGE="JavaScript">
 <!-- 

 function del(){                   
 if(window.confirm("ȷʵҪɾ����<%=trim(rs("Sub_name"))%>����")){            
  window.location = "Subject.asp?act=d&id=<%=rs("Sub_id")%>"; 
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
