<!--#include virtual="head.asp"-->
<!---->
<div id="bottom">
<div id="leftmenu">
<ul>
<li class="l" onclick="javascript:location.href='Company.asp'">结算单位</li>
<li class="l" onclick="javascript:location.href='Manager.asp'">项目经理</li>
<li class="l" onclick="javascript:location.href='Subject.asp'">费用项目</li>
<li class="l" onclick="javascript:location.href='Class.asp'">类　　别</li>
<li class="l1" onclick="javascript:location.href='Income.asp'">收入方式</li>
</ul>
</div>
<div id="rightContent">
<div id="rightContentadd">
<%

act=request("act")
id=request("id")
Inc_name=left(trim(request("Inc_name")),10)
if act="add" then
'验证用户名是否存在

	sql="select Inc_name from Income where Inc_name='"&Inc_name&"'"  ' 查询数据库中是否有重复记录
	
	set rs = conn.execute(sql)

	If not(rs.Eof And rs.Bof) Then
	 Response.Write ("<script language='javascript'>alert('“"&Inc_name&"”此名称已经存在，请修改重试');history.back(-1);</script>") ' 返回结果并进行编码转义
	response.end()
	end if

'添加
 set news=server.CreateObject("adodb.recordset")
     sql="select * from Income "
	 news.open sql,conn,3,3	 
	 news.addnew	 
	 news("Inc_name")=Inc_name
news("Inc_time")=now()
	 news("Inc_OpeID")=request.cookies("wuliuid") '用户id
     news.update
	 news.close
	 set news=nothing
response.write "<script language='javascript'>alert('提交成功');location.href('Income.asp');</script>"
end if
'修改
if act="mod" then
'验证用户名是否存在
	sql="select Inc_name from Income where Inc_name='"&Inc_name&"'and Inc_id<>"&id  ' 查询数据库中是否有重复记录
	
	set rs = conn.execute(sql)
	
	If not(rs.Eof And rs.Bof) Then
	 Response.Write ("<script language='javascript'>alert('“"&Inc_name&"”此名称已经存在，请修改重试');history.back(-1);</script>") ' 返回结果并进行编码转义
	response.end()
	end if
'
 set news=server.CreateObject("adodb.recordset")
     sql="select * from Income where Inc_id="&id

	 news.open sql,conn,3,3	
	  
	' news.addnew	 
	 news("Inc_name")=Cla_name
news("Inc_time")=now()
	 news("Inc_OpeID")=request.cookies("wuliuid") '用户id
     news.update
	 news.close
	 set news=nothing
response.write "<script language='javascript'>alert('修改成功');location.href('Income.asp');</script>"
end if

'删除数据
if act="d" then

dsql="delete from Income where Inc_id="&id
response.write dsql
conn.execute dsql
response.redirect"Income.asp"
end if%>




<%if act="m" and id<>""  then%>
<%set modrs=server.createobject("adodb.recordset")
modsql="select * from Income where Inc_Id="&id

modrs.open modsql,conn,3,3%>
<form name="wuliuform" method="post" action="?id=<%=id%>" onSubmit="return checkform();">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="181" height="65">&nbsp;</td>
      <td width="244">收入方式 <input name="Inc_name" type="text" id="Inc_name" size="20" value="<%=trim(modrs("Inc_name"))%>"/><%Old_Inc_name=trim(modrs("Inc_name"))%></td>
      <td width="245"><p></p></td>
      <td width="200"><input type="submit" name="button" id="button" value="修改收入方式" /><input type="hidden" name="act" value="mod" /></td>
    </tr>
  </table>
</form>
<%else%>
<form name="wuliuform" method="post" action="" onSubmit="return checkform();">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="181" height="65">&nbsp;</td>
      <td width="244">收入方式 <input name="Inc_name" type="text" id="Inc_name" size="20" /></td>
      <td width="245"><p></p></td>
      <td width="200"><input type="submit" name="button" id="button" value="新增收入方式" /><input type="hidden" name="act" value="add" /></td>
    </tr>
  </table>
</form>
<%end if%>
<script>
//提交验证
function checkform()
{
     var reg = /[^\w\u4e00-\u9fa5]/g;    // \w代表“数字、字母（不分大小写）、下划线”，\u4e00-\u9fa5代表汉字。 
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

//过程验证
window.onload=function(){
  var aInput = document.getElementsByTagName('input');
  var oName = aInput[0];
  var aP = document.getElementsByTagName('p');
  var name_msg = aP[0];
   var name_length = 0;
//会员名

  oName.onfocus = function(){
    name_msg.style.display = "inline";
    name_msg.innerHTML = "<i class='info'></i>推荐使用中文名";
  }

  oName.onblur = function(){

    //含有非法字符            
    var reg = /[^\w\u4e00-\u9fa5]/g;    // \w代表“数字、字母（不分大小写）、下划线”，\u4e00-\u9fa5代表汉字。 


    //不能为空
    if (this.value==""){
	
      name_msg.innerHTML = "<i class='no'></i>不能为空！";
    }
    else if(reg.test(this.value)){
      name_msg.innerHTML = '<i class="no"></i>含有非法字符！';
    }

    //OK
    else {
      name_msg.innerHTML = "<i class='yes'></i>OK！";

    }
  }

  }


</script>
</div>
<div id="rightContentlist">
<table style="width:100%">
<tr><th>系统ID</th>
<th>类别</th>
<th>添加时间</th><th>操作</th></tr>
<%  
'开始分页

dim intPage,page,pre,last,filepath 
'打开数据库  
set rs=server.createobject("adodb.recordset")
sql="select * from Income order by Inc_OrderId desc,Inc_id"
rs.PageSize = 100 '这里设定每页显示的记录数
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
do while not rs.eof and count<rs.PageSize
%> 
<tr onmousemove="changeTrColor(this)"><td><%=rs("Inc_id")%></td><td><%=rs("Inc_name")%></td><td><%=formatdatetime(rs("Inc_time"),2)%></td><td><a href="?act=m&id=<%=rs("Inc_id")%>"><img src="images/m.gif" border=0/></a>　<a href="javascript:del()"><img src="images/d.gif" border=0/></a><SCRIPT LANGUAGE="JavaScript">
 <!-- 

 function del(){                   
 if(window.confirm("确实要删除“<%=trim(rs("Inc_name"))%>”吗？")){            
  window.location = "Income.asp?act=d&id=<%=rs("Inc_id")%>"; 
  //提交的url         
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
