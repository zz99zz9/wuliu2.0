	<div id="mb"><table width="100%" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td width="6"><img src="images/ll.gif" width="6" height="37" /></td>
    <!--<td width="60" align="center"><a href="#" onClick="javascript:ShowDiv('sradd2')"><img src="images/l1.gif" width="36" height="26" /></a></td>
    <td width="5"><img src="images/lll.gif" width="2" height="37" /></td>
    
    <td width="60" align="center"><a href="#" onClick="javascript:HiddenDiv('sradd2')"><img src="images/l4.gif" width="36" height="26" /></a></td>
    <td width="5"><img src="images/lll.gif" width="2" height="37" /></td>-->
    <td width="60" align="center"><a href="Out_Chart.asp?<%=urlload%>" target="_blank"><img src="images/l3.gif" width="36" height="33" /></a></td>
    <td width="5"><img src="images/lll.gif" width="2" height="37" /></td>
    <td height="60"><!--#include virtual="inc/searchbar.asp"--></td>
    </tr>
</table>
</div>


<div>
    <% Server.ScriptTimeOut=950 %>
  <%'开始分页


  '打开数据库  
  set rs=server.createobject("adodb.recordset")
  if search="yes" then
  sql="select * from Exhibition "&sql1&" order by Exh_id desc"
  else
  sql="select top 30 * from Exhibition order by Exh_id desc"
  end if
  rs.PageSize = 10000 '这里设定每页显示的记录数
'response.write sql

  rs.open sql,conn,3,3
  %>

  <table width="980" border="0" cellspacing="1" cellpadding="0" class="datalist">
    <tr>
      <th>ID</th>
      <th>项目编号</th>
      <th>展会名称</th>
      <th>项目主管</th>
      <th>应收</th>
      <th>已收</th>
      <th>应付</th>
      <th>已付</th>
      <th>预计利润</th>
      <th>目前利润</th>
      <th>未收款</th>
    </tr>
    <%  

if err.number<>0 then
				response.write "数据库中暂时无数据"
				end if
				if rs.eof And rs.bof then
       				Response.Write "<p align='center' > 对不起，没有查询到您需要的信息！</p>"
   				else


i=0
do while not rs.eof 
i=i+1
%> 
    <tr onmousemove="changeTrColor(this)">
      <td><%=i%></td>
      <td><%=rs("Exh_Code")%></td>
      <td><%=rs("Exh_name")%></td>
      <td><%call Show_Supervisor_name(int(rs("Exh_Supid")))%></td>
      <td><%call Revenue_sum1(int(rs("Exh_id")))%></td>
<td><%call Revenue_sum2(int(rs("Exh_id")))%></td>
<td><%call Expense_sum1(int(rs("Exh_id")))%></td>
<td><%call Expense_sum2(int(rs("Exh_id")))%></td>
      <td><%call yjlr(int(rs("Exh_id")))%></td>
     <td><%call mqlr(int(rs("Exh_id")))%></td>
<td><%call wsk(int(rs("Exh_id")))%></td>
    </tr>
    <%
'	 Revenue_sum111=Revenue_sum111+int(Revenue_sum11)
'	 Revenue_sum222=Revenue_sum222+int(Revenue_sum22)
'	 Expense_sum111=Expense_sum111+int(Expense_sum11)
'	 Expense_sum222=Expense_sum222+int(Expense_sum22)
'	 yjlrhz11=yjlrhz11+int(yjlrhz)
'	 mqlr11=mqlr11+int(mqlrhz)
'	 wsk11=wsk11+int(wskhz)
 rs.movenext

 count=count+1
 loop
 
end if
rs.close
set rs=nothing
%> 
<!--    <tr >
      <td></td>
      <td></td>
      <td></td>
      <td><%=Revenue_sum111%></td>
<td><%=Revenue_sum222%></td>
<td><%=Expense_sum111%></td>
<td><%=Expense_sum222%></td>
      <td><%=yjlrhz11%></td>
     <td><%=mqlr11%></td>
<td><%=wsk11%></td>
    </tr>-->
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