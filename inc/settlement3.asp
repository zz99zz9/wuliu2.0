	<div id="mb"><table width="100%" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td width="6"><img src="images/ll.gif" width="6" height="37" /></td>
    <!--<td width="60" align="center"><a href="#" onClick="javascript:ShowDiv('sradd2')"><img src="images/l1.gif" width="36" height="26" /></a></td>
    <td width="5"><img src="images/lll.gif" width="2" height="37" /></td>
    
    <td width="60" align="center"><a href="#" onClick="javascript:HiddenDiv('sradd2')"><img src="images/l4.gif" width="36" height="26" /></a></td>
    <td width="5"><img src="images/lll.gif" width="2" height="37" /></td>-->
    <td width="60" align="center"><a href="Out_Settlement.asp?<%=urlload%>" target="_blank"><img src="images/l3.gif" width="36" height="33" /></a></td>
    <td width="5"><img src="images/lll.gif" width="2" height="37" /></td>
    <td height="60"> </td>
    </tr>
</table>
</div>


<div>
  <table width="980" border="0" cellspacing="1" cellpadding="0" class="datalist">
    <tr>
      <th>��Ŀ���</th>
      <th>չ������</th>
      <th>��Ŀ����</th>
      <th>Ӧ��</th>
      <th>����</th>
      <th>Ӧ��</th>
      <th>�Ѹ�</th>
      <th>Ԥ������</th>
      <th>Ŀǰ����</th>
      <th>δ�տ�</th>
    </tr>
    <%  
'��ʼ��ҳ


'�����ݿ�  
set rs=server.createobject("adodb.recordset")
sql="select * from Exhibition where Exh_id="&int(Exh_id)&" order by Exh_id desc"
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
    <tr onmousemove="changeTrColor(this)">
      <td><%=rs("Exh_Code")%></td>
      <td><%=rs("Exh_name")%></td>
      <td><%call Show_Supervisor_name(int(rs("Exh_Supid")))%></td>
      <td><%=Revenue_sum1(int(rs("Exh_id")))%></td>
<td><%=Revenue_sum2(int(rs("Exh_id")))%></td>
<td><%=Expense_sum1(int(rs("Exh_id")))%></td>
<td><%=Expense_sum2(int(rs("Exh_id")))%></td>
      <td><%=yjlr(int(rs("Exh_id")))%></td>
     <td><%=mqlr(int(rs("Exh_id")))%></td>
<td><%=wsk(int(rs("Exh_id")))%></td>
    </tr>
    <%
	 Revenue_sum111=Revenue_sum111+int(Revenue_sum11)
	 Revenue_sum222=Revenue_sum222+int(Revenue_sum22)
	 Expense_sum111=Expense_sum111+int(Expense_sum11)
	 Expense_sum222=Expense_sum222+int(Expense_sum22)
	 yjlrhz11=yjlrhz11+int(yjlrhz)
	 mqlr11=mqlr11+int(mqlrhz)
	 wsk11=wsk11+int(wskhz)
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