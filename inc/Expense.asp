      	<div id="mb"><table width="100%" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td width="6"><img src="images/ll.gif" width="6" height="37" /></td>
     <%if request.cookies("wuliuv")=0 then%>
    <td width="60" align="center"><a href="#" onClick="javascript:ShowDiv('sradd1')"><img src="images/l1.gif" width="36" height="26" /></a></td>
    <td width="5"><img src="images/lll.gif" width="2" height="37" /></td>
    
    <td width="60" align="center"><a href="#" onClick="javascript:HiddenDiv('sradd1')"><img src="images/l4.gif" width="36" height="26" /></a><!--<a href="#"><img src="images/l2.gif" width="37" height="31" /></a>--></td>
    <td width="5"><img src="images/lll.gif" width="2" height="37" /></td><%end if%>
    <td width="60" align="center"><a href="Out_Expense.asp?<%=urlload%>" target="_blank"><img src="images/l3.gif" width="36" height="33" /></a></td>
    <td height="60">&nbsp;</td>
    </tr>
</table>
</div>
<%

act=request("act")
id=request("id")
Exp_customer=trim(request("Exp_customer"))
Exp_Invoiceid=trim(request("Exp_Invoiceid"))
if Exp_Invoiceid="发票号" then Exp_Invoiceid=0
Exp_project=trim(request("Exp_project"))
Exp_amount1=trim(request("Exp_amount1"))
if Exp_amount1="应付金额" then Exp_amount1=0
Exp_amount2=trim(request("Exp_amount2"))
if Exp_amount2="已付金额" then Exp_amount2=0
Exp_Invoicename=trim(request("Exp_Invoicename"))
Exp_mode=trim(request("Exp_mode"))
Exp_content=trim(request("Exp_content"))
if act="add1" then

'验证用户名是否存在

sql="select * from Customer where Cus_name='"&Exp_customer&"'"  ' 查询数据库中是否有重复记录
	set rs = conn.execute(sql)
	If not(rs.Eof And rs.Bof) Then
	Exp_customer=rs("Cus_id")
		if Exp_Invoicename="" or Exp_Invoicename="发票抬头" then
			Exp_Invoicename=rs("Cus_name2") '企业名
		end if
	else
 Response.Write ("<script language='javascript'>alert('“"&Exp_customer&"”企业名称出错，请重新输入');history.back(-1);</script>") ' 返回结果并进行编码转义
	response.end()

	end if
	
	if 111=333 then
	sql="select Exp_Invoiceid from Expense where Exp_Invoiceid='"&Exp_Invoiceid&"'"  ' 查询数据库中是否有重复记录
	
	set rs = conn.execute(sql)

	If not(rs.Eof And rs.Bof) and Exp_Invoiceid<>0  Then
	 Response.Write ("<script language='javascript'>alert('“"&Exp_Invoiceid&"”此发票代码已经存在，请误重复提交');history.back(-1);</script>") ' 返回结果并进行编码转义
	response.end()
	end if
end if
'添加
 set news=server.CreateObject("adodb.recordset")
     sql="select * from Expense "
	 news.open sql,conn,3,3	 
	 news.addnew	 
	 news("Exp_customer")=Exp_customer
	 news("Exp_Invoiceid")=Exp_Invoiceid
	 news("Exp_project")=Exp_project
	 news("Exp_amount1")=Exp_amount1
 news("Exp_amount2")=Exp_amount2
	 news("Exp_Invoicename")=Exp_Invoicename
	 news("Exp_mode")=Exp_mode
	 news("Exp_content")=Exp_content
	 news("Exp_time")=now()
	 news("Exp_OpeID")=request.cookies("wuliuid") '用户id
	 news("Exp_Exhid")=Exh_id
     news.update
	 news.close
	 set news=nothing
response.write "<script language='javascript'>alert('提交成功!');location.href('index.asp?Riframe=1&"&urlload&"');</script>"
end if
'修改

if act="mod1" then
'验证用户名是否存在

sql="select * from Customer where Cus_name='"&Exp_customer&"'"  ' 查询数据库中是否有重复记录
	set rs = conn.execute(sql)
	If not(rs.Eof And rs.Bof) Then
	Exp_customer=rs("Cus_id")
		if Exp_Invoicename="" or Exp_Invoicename="发票抬头" then
			Exp_Invoicename=rs("Cus_name2") '企业名
		end if
	else
 Response.Write ("<script language='javascript'>alert('“"&Exp_customer&"”企业名称出错，请重新输入');history.back(-1);</script>") ' 返回结果并进行编码转义
	response.end()

	end if

'
 set news=server.CreateObject("adodb.recordset")
     sql="select * from Expense where Exp_id="&id

	 news.open sql,conn,3,3	
	  
	' news.addnew	 
	 news("Exp_customer")=Exp_customer
	 news("Exp_Invoiceid")=Exp_Invoiceid
	 news("Exp_project")=Exp_project
	 news("Exp_amount1")=Exp_amount1
 news("Exp_amount2")=Exp_amount2
	 news("Exp_Invoicename")=Exp_Invoicename
	 news("Exp_mode")=Exp_mode
	 news("Exp_content")=Exp_content
	 news("Exp_time")=now()
	 news("Exp_OpeID")=request.cookies("wuliuid") '用户id
	 news("Exp_Exhid")=Exh_id
     news.update
	 news.close
	 set news=nothing
response.write "<script language='javascript'>alert('修改成功');location.href('index.asp?Riframe=1&"&urlload&"');</script>"
end if

'删除数据
if act="d1" then

dsql="delete from Expense where Exp_id="&id
response.write dsql
conn.execute dsql
response.redirect"index.asp?"&urlload&"&Riframe=1"
end if%>
<div id="sradd1">
<form id="wuliuform" name="wuliuform" method="post" action="">
  <table width="980" border="0" cellspacing="1" cellpadding="0" class="datalist">
    <tr>
      <td width="211">
        <p><input name="Exp_customer" type="text" id="Exp_customer" value="客户名称" size="15" onfocus="javascript:if(this.value=='客户名称')this.value='';" onblur="if(this.value==''){this.value='客户名称'}" onkeyup="searchSuggest2();" AUTOCOMPLETE="off" class="tete"/><br /><div class="search_suggest" style="display:none;"></div>
        <!---->
         
        </p>
        <p>
                    <select name="Exp_project" id="Exp_project">
          <option value="0">费用项目</option>
          
          <%set rs=server.createobject("adodb.recordset")
sql="select * from Subject order by Sub_OrderId desc,Sub_id"


rs.open sql,conn,3,3
do while not rs.eof%>
<option value="<%=rs("Sub_id")%>" <%if rs("Sub_id")=2 then%> selected="selected"<%end if%>><%=rs("Sub_name")%></option>
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
        <input name="Exp_amount1" type="text" id="Exp_amount1" value="应付金额" size="10" onfocus="javascript:if(this.value=='应付金额')this.value='';" onblur="if(this.value==''){this.value='应付金额'}"/>
      </p>
        <p><input name="Exp_amount2" type="text" id="Exp_amount2" value="已付金额" size="10" onfocus="javascript:if(this.value=='已付金额')this.value='';" onblur="if(this.value==''){this.value='已付金额'}"/></p></td>
      <td width="190"><p>
        <input name="Exp_Invoiceid" type="text" id="Exp_Invoiceid" value="发票号" onfocus="javascript:if(this.value=='发票号')this.value='';" onblur="if(this.value==''){this.value='发票号'}"/>
      </p>
        <p>        
          <input name="Exp_Invoicename" type="text" id="Exp_Invoicename" value="发票抬头" onfocus="javascript:if(this.value=='发票抬头')this.value='';" onblur="if(this.value==''){this.value='发票抬头'}"/>
        </p></td>
      <td width="211"><p>
                           <select name="Exp_mode" id="Exp_mode">
          <option value="0">支出方式</option>
          
          <%set rs=server.createobject("adodb.recordset")
sql="select * from Income order by Inc_id"


rs.open sql,conn,3,3
do while not rs.eof%>
<option value="<%=rs("Inc_id")%>" <%if rs("Inc_id")=2 then%> selected="selected"<%end if%>><%=rs("Inc_name")%></option>
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
          <input name="Exp_content" type="text" id="Exp_content" value="备注" size="20" onfocus="javascript:if(this.value=='备注')this.value='';" onblur="if(this.value==''){this.value='备注'}" />
        </p></td>
      <td width="106"><input type="submit" name="button" id="button" value=" 保 存 " /><input type="hidden" name=act id=act value="add1" /></td>
      </tr>
  </table></form>
</div>
<!--修改信息-->
<% if act="m1" then%>
<div id="srmod1">
<%set rs1=server.createobject("adodb.recordset")
sql1="select * from Expense where Exp_id="&id&" order by Exp_id"
rs1.open sql1,conn,3,3%>
<form id="wuliuform" name="wuliuform" method="post" action="?act=mod1&id=<%=id%>&<%=urlload%>">
  <table width="980" border="0" cellspacing="1" cellpadding="0" class="datalist">
    <tr>
      <td width="211">
      <p><input name="Exp_customer" type="text" id="Exp_customer2" value="<%call Show_customer_name(rs1("Exp_customer"))%>" size="15" onfocus="javascript:if(this.value=='客户名称')this.value='';" onblur="if(this.value==''){this.value='客户名称'}" onkeyup="searchSuggest4();" AUTOCOMPLETE="off" class="tete"/><br /><div class="search_suggest" style="display:none;"></div>
        <!---->
         
        </p>
       
        <p>
                    <select name="Exp_project" id="Exp_project">
          <option value="0">费用项目</option>
          
          <%set rs=server.createobject("adodb.recordset")
sql="select * from Subject order by Sub_OrderId desc,Sub_id"


rs.open sql,conn,3,3
do while not rs.eof%>
<option value="<%=rs("Sub_id")%>" <%if int(rs1("Exp_project"))=int(rs("Sub_id")) then%> selected="selected"<%end if%>><%=rs("Sub_name")%></option>
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
        <input name="Exp_amount1" type="text" id="Exp_amount1" value="<%=rs1("Exp_amount1")%>" size="10"/>
      </p>
        <p><input name="Exp_amount2" type="text" id="Exp_amount2" value="<%=rs1("Exp_amount2")%>" size="10"/></p></td>
      <td width="190"><p>
        <input name="Exp_Invoiceid" type="text" id="Exp_Invoiceid" value="<%=rs1("Exp_Invoiceid")%>"/>
      </p>
        <p>        
          <input name="Exp_Invoicename" type="text" id="Exp_Invoicename" value="<%=rs1("Exp_Invoicename")%>"/>
        </p></td>
      <td width="211"><p>
                           <select name="Exp_mode" id="Exp_mode">
          <option  value="0">支出方式</option>
          
          <%set rs=server.createobject("adodb.recordset")
sql="select * from Income order by Inc_id"


rs.open sql,conn,3,3
do while not rs.eof%>
<option value="<%=rs("Inc_id")%>" <%if int(rs1("Exp_mode"))=int(rs("Inc_id")) then%> selected="selected"<%end if%>><%=rs("Inc_name")%></option>
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
          <input name="Exp_content" type="text" id="Exp_content" value="备注" size="20" onfocus="javascript:if(this.value=='备注')this.value='';" onblur="if(this.value==''){this.value='备注'}" />
        </p></td>
      <td width="106"><input type="submit" name="button" id="button" value=" 保 存 " /><input type="hidden" name=act id=act value="mod" /></td>
      </tr>
  </table></form>
</div>
<%end if%>
<div>
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
      <th>操作</th>
    </tr>
    <%  
'开始分页


'打开数据库  
set rs=server.createobject("adodb.recordset")
sql="select * from Expense where Exp_Exhid="&int(Exh_id)&" order by Exp_id desc"


rs.PageSize = 200 '这里设定每页显示的记录数
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
      <td> <%if request.cookies("wuliuv")=0 then%><a href="?Riframe=1&<%=urlload%>&act=m1&amp;id=<%=rs("Exp_id")%>"><img src="images/m.gif" border="0"/></a>　<a href="javascript:del<%=rs("Exp_id")%>()"><img src="images/d.gif" border="0"/></a><SCRIPT LANGUAGE="JavaScript">
 <!-- 

 function del<%=rs("Exp_id")%>(){                   
 if(window.confirm("确实要删除“<%call Show_customer_name(int(rs("Exp_customer")))%>”的付款信息吗？")){            
  window.location = "index.asp?act=d1&id=<%=rs("Exp_id")%>&<%=urlload%>"; 
  //提交的url         
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