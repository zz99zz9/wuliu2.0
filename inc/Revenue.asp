      	<div id="mb"><table width="100%" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td width="6"><img src="images/ll.gif" width="6" height="37" /></td> <%if request.cookies("wuliuv")=0 then%>
    <td width="60" align="center"><a href="#" onClick="javascript:ShowDiv('sradd')"><img src="images/l1.gif" width="36" height="26" /></a></td>
    <td width="5"><img src="images/lll.gif" width="2" height="37" /></td>
    
    <td width="60" align="center"><a href="#" onClick="javascript:HiddenDiv('sradd')"><img src="images/l4.gif" width="36" height="26" /></a><!--<a href="#"><img src="images/l2.gif" width="37" height="31" /></a>--></td>
    <td width="5"><img src="images/lll.gif" width="2" height="37" /></td><%end if%>
    <td width="60" align="center"><a href="Out_Revenue.asp?<%=urlload%>" target="_blank"><img src="images/l3.gif" width="36" height="33" /></a></td>
    <td height="60">&nbsp;</td>
    </tr>
</table>
</div>
<%

act=request("act")
id=request("id")
Rev_customer=trim(request("Rev_customer"))

searchkey=trim(request("searchkey"))
if searchkey<>"" then
sql="and searchkey"
end if

Rev_Invoiceid=trim(request("Rev_Invoiceid"))
if Rev_Invoiceid="��Ʊ��" then Rev_Invoiceid=0

Rev_project=trim(request("Rev_project"))
'if Rev_project="" then Rev_project=0
Rev_amount1=trim(request("Rev_amount1"))
if Rev_amount1="Ӧ�ս��" then Rev_amount1=0
Rev_amount2=trim(request("Rev_amount2"))
if Rev_amount2="���ս��" then Rev_amount2=0
Rev_Invoicename=trim(request("Rev_Invoicename"))
Rev_mode=trim(request("Rev_mode"))

Rev_content=trim(request("Rev_content"))
'��֤������Ϊ��

if act="add" then
'��֤��ҵ����Ƿ����
	sql="select * from Customer where Cus_name='"&Rev_customer&"'"  ' ��ѯ���ݿ����Ƿ����ظ���¼
	set rs = conn.execute(sql)
	If not(rs.Eof And rs.Bof) Then
	Rev_customer=rs("Cus_id")
			if Rev_Invoicename="" or Rev_Invoicename="��Ʊ̧ͷ" then
			Rev_Invoicename=rs("Cus_name2") '��ҵ��
		end if

	else
 Response.Write ("<script language='javascript'>alert('��"&Rev_customer&"����ҵ���Ƴ�������������');history.back(-1);</script>") ' ���ؽ�������б���ת��
	response.end()

	end if
	

if 222=333 then	
	'�жϷ�Ʊ���Ƿ��ظ�
	sql="select Rev_Invoiceid from Revenue where Rev_Invoiceid='"&Rev_Invoiceid&"'"  ' ��ѯ���ݿ����Ƿ����ظ���¼
	
	set rs = conn.execute(sql)

	If not(rs.Eof And rs.Bof) and Rev_Invoiceid<>0 Then
	 Response.Write ("<script language='javascript'>alert('��"&Rev_Invoiceid&"���˷�Ʊ�����Ѿ����ڣ������ظ��ύ');history.back(-1);</script>") ' ���ؽ�������б���ת��
	response.end()
	end if
end if
'���
 set news=server.CreateObject("adodb.recordset")
     sql="select * from Revenue "
	 news.open sql,conn,3,3	 
	 news.addnew	 
	 news("Rev_customer")=Rev_customer
	 news("Rev_Invoiceid")=Rev_Invoiceid
	 news("Rev_project")=Rev_project
	 news("Rev_amount1")=Rev_amount1
	 news("Rev_amount2")=Rev_amount2
	 news("Rev_Invoicename")=Rev_Invoicename
	 news("Rev_mode")=Rev_mode
	 news("Rev_content")=Rev_content
	 news("Rev_time")=now()
	 news("Rev_OpeID")=request.cookies("wuliuid") '�û�id
	 news("Rev_Exhid")=Exh_id
     news.update
	 news.close
	 set news=nothing
response.write "<script language='javascript'>alert('�ύ�ɹ�');location.href('index.asp?"&urlload&"');</script>"
end if
'�޸�

if act="mod" then
'��֤��ҵ����Ƿ����
	sql="select * from Customer where Cus_name='"&Rev_customer&"'"  ' ��ѯ���ݿ����Ƿ����ظ���¼
	set rs = conn.execute(sql)
	If not(rs.Eof And rs.Bof) Then
	Rev_customer=rs("Cus_id")
			if Rev_Invoicename="" or Rev_Invoicename="��Ʊ̧ͷ" then
			Rev_Invoicename=rs("Cus_name2") '��ҵ��
		end if

	else
 Response.Write ("<script language='javascript'>alert('��"&Rev_customer&"����ҵ���Ƴ�������������');history.back(-1);</script>") ' ���ؽ�������б���ת��
	response.end()

	end if

'
 set news=server.CreateObject("adodb.recordset")
     sql="select * from Revenue where Rev_id="&id

	 news.open sql,conn,3,3	
	  
	' news.addnew	 
	 news("Rev_customer")=Rev_customer
	 news("Rev_Invoiceid")=Rev_Invoiceid
	 news("Rev_project")=Rev_project
	 news("Rev_amount1")=Rev_amount1
	 news("Rev_amount2")=Rev_amount2
	 news("Rev_Invoicename")=Rev_Invoicename
	 news("Rev_mode")=Rev_mode
	 news("Rev_content")=Rev_content
	 news("Rev_time")=now()
	 news("Rev_OpeID")=request.cookies("wuliuid") '�û�id
	 news("Rev_Exhid")=Exh_id
     news.update
	 news.close
	 set news=nothing
response.write "<script language='javascript'>alert('�޸ĳɹ�');location.href('index.asp?"&urlload&"');</script>"
end if

'ɾ������
if act="d" then

dsql="delete from Revenue where Rev_id="&id
response.write dsql
conn.execute dsql
response.redirect"index.asp?"&urlload  'Exh_code="&Exh_code&"&ECount="&ECount&"&Exh_id="&Exh_id
end if%>
<div id="sradd">
<form id="wuliuform" name="wuliuform" method="post" action="">
  <table width="980" border="0" cellspacing="1" cellpadding="0" class="datalist">
    <tr>
      <td width="211">
        <p>
        <input name="Rev_customer" type="text" id="Rev_customer2" value="�ͻ�����" size="15" onfocus="javascript:if(this.value=='�ͻ�����')this.value='';" onblur="if(this.value==''){this.value='�ͻ�����'}" onkeyup="searchSuggest3();" AUTOCOMPLETE="off" class="tete"/><br /><div class="search_suggest" style="display:none"></div>
          
        </p>
        <p>
                    <select name="Rev_project" id="Rev_project">
                      <option value="0">������Ŀ</option>
                   
          
          <%set rs=server.createobject("adodb.recordset")
sql="select * from Subject order by Sub_OrderId desc,Sub_id"


rs.open sql,conn,3,3
do while not rs.eof%>   <option value="<%=rs("Sub_id")%>" <%if rs("Sub_id")=2 then%> selected="selected"<%end if%>><%=rs("Sub_name")%></option>
            <%
 rs.movenext
 count=count+1
 loop
rs.close
set rs=nothing
%> 
        </select>
        </p>
      </td>
      <td width="154"><p>
        <input name="Rev_amount1" type="text" id="Rev_amount1" value="Ӧ�ս��" size="10" onfocus="javascript:if(this.value=='Ӧ�ս��')this.value='';" onblur="if(this.value==''){this.value='Ӧ�ս��'}" /><!--onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"-->
      </p>
        <p>
          <input name="Rev_amount2" type="text" id="Rev_amount2" value="���ս��" size="10" onfocus="javascript:if(this.value=='���ս��')this.value='';" onblur="if(this.value==''){this.value='���ս��'}"/><!-- onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"-->
        </p></td>
      <td width="190"><p>
        <input name="Rev_Invoiceid" type="text" id="Rev_Invoiceid" value="��Ʊ��" onfocus="javascript:if(this.value=='��Ʊ��')this.value='';" onblur="if(this.value==''){this.value='��Ʊ��'}"/><!-- onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"-->
      </p>
        <p>        
          <input name="Rev_Invoicename" type="text" id="Rev_Invoicename" value="��Ʊ̧ͷ" onfocus="javascript:if(this.value=='��Ʊ̧ͷ')this.value='';" onblur="if(this.value==''){this.value='��Ʊ̧ͷ'}"/>
        </p></td>
      <td width="211"><p>
                           <select name="Rev_mode" id="Rev_mode">
                             <option value="0" selected="selected">���뷽ʽ</option>
                            
          
          <%set rs=server.createobject("adodb.recordset")
sql="select * from Income order by Inc_id"


rs.open sql,conn,3,3
do while not rs.eof%> <option value="<%=rs("Inc_id")%>" <%if rs("Inc_id")=2 then%>selected="selected"<%end if%>><%=rs("Inc_name")%></option>
            <%
 rs.movenext
 count=count+1
 loop
rs.close
set rs=nothing
%> 
        </select>
        </p>
        <p>
          <input name="Rev_content" type="text" id="Rev_content" value="��ע" size="20" onfocus="javascript:if(this.value=='��ע')this.value='';" onblur="if(this.value==''){this.value='��ע'}" />
        </p></td>
      <td width="106"><input type="submit" name="button" id="button" value=" �� �� " /><input type="hidden" name=act id=act value="add" /></td>
      </tr>
  </table></form>
</div>
<!--�޸���Ϣ-->
<% if act="m" then%>
<div id="srmod">
<%set rs1=server.createobject("adodb.recordset")
sql1="select * from Revenue where Rev_id="&id&" order by Rev_id"
rs1.open sql1,conn,3,3%>
<form id="wuliuform" name="wuliuform" method="post" action="?act=mod&id=<%=id%>&<%=urlload%>">
  <table width="980" border="0" cellspacing="1" cellpadding="0" class="datalist">
    <tr>
      <td width="211">
        <p>
          
          <input name="Rev_customer" type="text" id="Rev_customer" value="<%call Show_customer_name(trim(rs1("Rev_customer")))%>" size="15" onfocus="javascript:if(this.value=='�ͻ�����')this.value='';" onblur="if(this.value==''){this.value='�ͻ�����'}" onkeyup="searchSuggest();" AUTOCOMPLETE="off" class="tete"/><br /><div class="search_suggest" style="display:none"></div>
          

        </p>
        <p>
                    <select name="Rev_project" id="Rev_project">
                      <option value="0" >������Ŀ</option>
                      
          
          <%set rs=server.createobject("adodb.recordset")
sql="select * from Subject order by Sub_OrderId desc,Sub_id"


rs.open sql,conn,3,3
do while not rs.eof%><option value="<%=rs("Sub_id")%>" <%if int(rs1("Rev_project"))=int(rs("Sub_id")) then%> selected="selected"<%end if%>><%=rs("Sub_name")%></option>
            <%
 rs.movenext
 count=count+1
 loop
rs.close
set rs=nothing
%> 
        </select>
        </p>
      </td>
      <td width="154"><p>
        <input name="Rev_amount1" type="text" id="Rev_amount1" value="<%=rs1("Rev_amount1")%>" size="10"/><!-- onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"-->
      </p>
        <p>
          <input name="Rev_amount2" type="text" id="Rev_amount2" value="<%=rs1("Rev_amount2")%>" size="10" /><!--onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"-->
        </p></td>
      <td width="190"><p>
        <input name="Rev_Invoiceid" type="text" id="Rev_Invoiceid" value="<%=rs1("Rev_Invoiceid")%>"/><!-- onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"-->
      </p>
        <p>        
          <input name="Rev_Invoicename" type="text" id="Rev_Invoicename" value="<%=trim(rs1("Rev_Invoicename"))%>"/>
        </p></td>
      <td width="211"><p>
                           <select name="Rev_mode" id="Rev_mode">
                             <option value="0" >���뷽ʽ</option>
                            
          
          <%set rs=server.createobject("adodb.recordset")
sql="select * from Income order by Inc_id"


rs.open sql,conn,3,3
do while not rs.eof%> <option value="<%=rs("Inc_id")%>" <%if int(rs1("Rev_mode"))=int(rs("Inc_id")) then%> selected="selected"<%end if%>><%=rs("Inc_name")%></option>
            <%
 rs.movenext
 count=count+1
 loop
rs.close
set rs=nothing
%> 
        </select>
        </p>
        <p>
          <input name="Rev_content" type="text" id="Rev_content" value="<%=trim(rs1("Rev_content"))%>" size="20"  />
        </p></td>
      <td width="106"><input type="submit" name="button" id="button" value=" �� �� " /><input type="hidden" name=act id=act value="mod" /></td>
      </tr>
  </table></form>
</div>
<%end if%>
<div>
  <table width="980" border="0" cellspacing="1" cellpadding="0" class="datalist">
    <tr>
      <th>���</th>
      <th>�ͻ�����</th>
      <th>������Ŀ</th>
      <th>Ӧ��</th>
      <th>����</th>
      <th>��Ʊ��</th>
      <th>��Ʊ̧ͷ</th>
      <th>���뷽ʽ</th>
      <th>��ע</th>
      <th>������</th>
      <th>��������</th>
      <th>����</th>
    </tr>
    <%  
'��ʼ��ҳ

dim intPage,page,pre,last,filepath 
'�����ݿ�  
set rs=server.createobject("adodb.recordset")
sql="select * from Revenue where Rev_Exhid="&int(Exh_id)&" order by Rev_id desc"

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
do while not rs.eof 
%> 
    <tr onmousemove="changeTrColor(this)" <%if rs("Rev_amount1")<=rs("Rev_amount2") and rs("Rev_amount1")<>0 then%> style="color:#ff0000;"<%end if%>>
      <td><%=Exh_code%></td>
      <td><%call Show_customer_name(int(rs("Rev_customer")))%></td>
      <td><%call Show_Subject_name(int(rs("Rev_project")))%></td>
      <td><%if rs("Rev_amount1")=0 then%>0.00<%else%><%=FormatNumber(rs("Rev_amount1"))%><%end if%></td>
      <td><%if rs("Rev_amount2")=0 then%>0.00<%else%><%=FormatNumber(rs("Rev_amount2"))%><%end if%></td>
      <td><%=rs("Rev_Invoiceid")%></td>
      <td><%=rs("Rev_Invoicename")%></td>
      <td><%call Show_Income_name(int(rs("Rev_mode")))%></td>
      <td><%=rs("Rev_content")%></td>
      <td><%call Show_operator_name(int(rs("Rev_Opeid")))%></td>
      <td><%=formatdatetime(rs("Rev_time"),2)%></td>
      <td> <%if request.cookies("wuliuv")=0 then%><a href="?<%=urlload%>&act=m&amp;id=<%=rs("Rev_id")%>"><img src="images/m.gif" border="0"/></a>��<a href="javascript:del<%=rs("Rev_id")%>()"><img src="images/d.gif" border="0"/></a><SCRIPT LANGUAGE="JavaScript">
 <!-- 

 function del<%=rs("Rev_id")%>(){                   
 if(window.confirm("ȷʵҪɾ����<%call Show_customer_name(int(rs("Rev_customer")))%>���ĸ�����Ϣ��")){            
  window.location = "index.asp?act=d&id=<%=rs("Rev_id")%>&<%=urlload%>"; 
  //�ύ��url         
  }else{             
  return;         
  }   
    } //--> 
    </SCRIPT><%end if%>
</td>
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
 <td>����</td>
      <td></td>
      <td></td>
      <td><b><%=Revenue_sum1(Exh_id)%></b></td>
<td><b><%=Revenue_sum2(Exh_id)%></b></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
</tr>
<%end if%>
  </table>
  <script type="text/javascript">
function changeTrColor(obj){    
    var _table=obj.parentNode;
    for (var i=0;i<_table.rows.length;i++){
        _table.rows[i].style.backgroundColor="";
    }    
    obj.style.backgroundColor="#FEE8EA";
}
</script>
</div>