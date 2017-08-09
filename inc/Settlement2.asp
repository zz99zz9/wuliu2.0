      	<div id="mb"><table width="100%" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td width="6"><img src="images/ll.gif" width="6" height="37" /></td>
    <td width="60" align="center"><a href="#" onClick="javascript:ShowDiv('sradd2')"><img src="images/l1.gif" width="36" height="26" /></a></td>
    <td width="5"><img src="images/lll.gif" width="2" height="37" /></td>
    
    <td width="60" align="center"><a href="#" onClick="javascript:HiddenDiv('sradd2')"><img src="images/l4.gif" width="36" height="26" /></a><!--<a href="#"><img src="images/l2.gif" width="37" height="31" /></a>--></td>
    <td width="5"><img src="images/lll.gif" width="2" height="37" /></td>
    <td width="60" align="center"><a href="Out_Settlement.asp?<%=urlload%>" target="_blank"><img src="images/l3.gif" width="36" height="33" /></a></td>
    <td height="60">&nbsp;</td>
    </tr>
</table>
</div>


<div>
  <table width="980" border="0" cellspacing="1" cellpadding="0" class="datalist">
    <tr>
      <th>流水号</th>
      <th>结算单位</th>
      <th>费用项目</th>
      <th>总价</th>
      <th>收款方式</th>
    </tr>
    <%  
'开始分页


'打开数据库  
set rs=server.createobject("adodb.recordset")
sql="select * from Expense where Exp_Exhid="&int(Exh_id)&" order by Exp_id desc"
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
do while not rs.eof 
%> 
    <tr onmousemove="changeTrColor(this)">
      <td><%=Exh_code%></td>
      <td><%call Show_customer_name(int(rs("Exp_customer")))%></td>
      <td><%call Show_Subject_name(int(rs("Exp_project")))%></td>
      <td><%=rs("Exp_amount")%></td>


      <td><%call Show_Income_name(int(rs("Exp_mode")))%></td>
     

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