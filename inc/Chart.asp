<%response.cookies("S_year")=request("S_year")
response.cookies("S_moon")=request("S_moon")
response.cookies("E_year")=request("E_year")
response.cookies("E_moon")=request("E_moon")

if timeClear="Clear" or request.cookies("S_moon")="" then
response.cookies("S_year")=2016
response.cookies("S_moon")=1
response.cookies("E_year")=year(now())
response.cookies("E_moon")=month(now())
end if
S_year=request.cookies("S_year")
S_moon=request.cookies("S_moon")
E_year=request.cookies("E_year")
E_moon=request.cookies("E_moon")%>
 <script language="javascript">
function checkform()


{

	if (document.wuliuform.S_year.value>document.wuliuform.E_year.value)
		{
			alert("起始年份不能超过终止年份！");
			//document.form1.title.focus();
			return false;
		}
		else if(document.wuliuform.S_year.value=document.wuliuform.E_year.value && document.wuliuform.S_moon.value>document.wuliuform.E_moon.value)
		{alert("起始月份不能超过终止月份！");
		return false;
			}else{
			
	return true;}
}
</script>	<div id="mb"><table width="100%" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td width="6"><img src="images/ll.gif" width="6" height="37" /></td>
    <!--<td width="60" align="center"><a href="#" onClick="javascript:ShowDiv('sradd2')"><img src="images/l1.gif" width="36" height="26" /></a></td>
    <td width="5"><img src="images/lll.gif" width="2" height="37" /></td>
    
    <td width="60" align="center"><a href="#" onClick="javascript:HiddenDiv('sradd2')"><img src="images/l4.gif" width="36" height="26" /></a></td>
    <td width="5"><img src="images/lll.gif" width="2" height="37" /></td>-->
    <td width="60" align="center"><a href="Out_Chart.asp?<%=urlload%>" target="_blank"><img src="images/l3.gif" width="36" height="33" /></a></td>
    <td width="5"><img src="images/lll.gif" width="2" height="37" /></td>
    <td height="60"> <form id="wuliuform" name="wuliuform" method="post" action="" onSubmit="return checkform();">
    　从　<select name="S_year" id="S_year">
            <option value="2016" <%if S_year=2016 then %>selected="selected"<%end if%>>2016</option>
            <option value="2017" <%if S_year=2017 then %>selected="selected"<%end if%>>2017</option>
            <option value="2018" <%if S_year=2018 then %>selected="selected"<%end if%>>2018</option>
            <option value="2019" <%if S_year=2019 then %>selected="selected"<%end if%>>2019</option>
            <option value="2020" <%if S_year=2020 then %>selected="selected"<%end if%>>2020</option>
            <option value="2021" <%if S_year=2021 then %>selected="selected"<%end if%>>2021</option>
            <option value="2022" <%if S_year=2022 then %>selected="selected"<%end if%>>2022</option>
          </select>
            <select name="S_moon" id="S_moon">
              <option value="1" <%if S_moon=1 then %>selected="selected"<%end if%>>1</option>
              <option value="2" <%if S_moon=2 then %>selected="selected"<%end if%>>2</option>
              <option value="3" <%if S_moon=3 then %>selected="selected"<%end if%>>3</option>
              <option value="4" <%if S_moon=4 then %>selected="selected"<%end if%>>4</option>
              <option value="5" <%if S_moon=5 then %>selected="selected"<%end if%>>5</option>
              <option value="6" <%if S_moon=6 then %>selected="selected"<%end if%>>6</option>
              <option value="7" <%if S_moon=7 then %>selected="selected"<%end if%>>7</option>
              <option value="8" <%if S_moon=8 then %>selected="selected"<%end if%>>8</option>
              <option value="9" <%if S_moon=9 then %>selected="selected"<%end if%>>9</option>
              <option value="10" <%if S_moon=10 then %>selected="selected"<%end if%>>10</option>
              <option value="11" <%if S_moon=11 then %>selected="selected"<%end if%>>11</option>
              <option value="12" <%if S_moon=12 then %>selected="selected"<%end if%>>12</option>
            </select>　到　<select name="E_year" id="E_year">
            <option value="2016" <%if E_year=2016 then %>selected="selected"<%end if%>>2016</option>
            <option value="2017" <%if E_year=2017 then %>selected="selected"<%end if%>>2017</option>
            <option value="2018" <%if E_year=2018 then %>selected="selected"<%end if%>>2018</option>
            <option value="2019" <%if E_year=2019 then %>selected="selected"<%end if%>>2019</option>
            <option value="2020" <%if E_year=2020 then %>selected="selected"<%end if%>>2020</option>
            <option value="2021" <%if E_year=2021 then %>selected="selected"<%end if%>>2021</option>
            <option value="2022" <%if E_year=2022 then %>selected="selected"<%end if%>>2022</option>
          </select>
        <select name="E_moon" id="E_moon">
              <option value="1" <%if E_moon=1 then %>selected="selected"<%end if%>>1</option>
              <option value="2" <%if E_moon=2 then %>selected="selected"<%end if%>>2</option>
              <option value="3" <%if E_moon=3 then %>selected="selected"<%end if%>>3</option>
              <option value="4" <%if E_moon=4 then %>selected="selected"<%end if%>>4</option>
              <option value="5" <%if E_moon=5 then %>selected="selected"<%end if%>>5</option>
              <option value="6" <%if E_moon=6 then %>selected="selected"<%end if%>>6</option>
              <option value="7" <%if E_moon=7 then %>selected="selected"<%end if%>>7</option>
              <option value="8" <%if E_moon=8 then %>selected="selected"<%end if%>>8</option>
              <option value="9" <%if E_moon=9 then %>selected="selected"<%end if%>>9</option>
              <option value="10" <%if E_moon=10 then %>selected="selected"<%end if%>>10</option>
              <option value="11" <%if E_moon=11 then %>selected="selected"<%end if%>>11</option>
              <option value="12" <%if E_moon=12 then %>selected="selected"<%end if%>>12</option>
            </select>
            <input type="submit" name="button" id="button" value="搜索" />　<a href="?<%=urlload%>&Riframe=4&time=Clear">查看全部</a></form></td>
    </tr>
</table>
</div>


<div>
  <table width="980" border="0" cellspacing="1" cellpadding="0" class="datalist">
    <tr>
      <th>时间</th>
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
'开始分页
sql1="where Exh_year>="&S_year&" and Exh_moon>="&S_moon&" and Exh_year<="&E_year&" and Exh_moon<="&E_moon&""
  if E_year>S_year then
sql1="where (Exh_year="&S_year&" and Exh_moon>="&S_moon&") or (Exh_year="&E_year&" and Exh_moon<="&E_moon&")"
  end if
'sql1="where Exh_year>="&S_year&" and Exh_moon>="&S_moon&" and Exh_year<="&E_year&" and Exh_moon<="&E_moon&""
'打开数据库  
set rs=server.createobject("adodb.recordset")
sql="select * from Exhibition "&sql1&" order by Exh_id desc"
rs.PageSize = 1000 '这里设定每页显示的记录数
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
      <td><%=rs("Exh_year")%>-<%=rs("Exh_moon")%></td>
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