
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<%response.cookies("S_year")=request("S_year")
response.cookies("S_moon")=request("S_moon")
response.cookies("E_year")=request("E_year")
response.cookies("E_moon")=request("E_moon")
response.cookies("Sup_id")=request("Sup_id")
skey=request("skey")
if skey="1" then search="yes" else search="no" end if
if timeClear="Clear" or request.cookies("S_moon")="" then
skey=""
response.cookies("S_year")=2016
response.cookies("S_moon")=1
response.cookies("E_year")=year(now())
response.cookies("E_moon")=month(now())
end if
S_year=request.cookies("S_year")
S_moon=request.cookies("S_moon")
E_year=request.cookies("E_year")
E_moon=request.cookies("E_moon")
if request.cookies("Sup_id")<>"" then
Sup_id=int(request.cookies("Sup_id"))
end if
  s_time=FormatNumber(S_year&"."&S_moon,2,False,False,False)
  e_time=FormatNumber(E_year&"."&E_moon,2,False,False,False)
  sql1="where w_time>="&s_time&" and w_time<="&e_time&""
  if Sup_id<>"0" then
    sql1=sql1+" and Supid="&Sup_id
  end if
%>
 <script language="javascript">
function checkform()

{

	if (document.wuliuform.S_year.value>document.wuliuform.E_year.value)
		{
			alert("起始年份不能超过终止年份！");
			//document.form1.title.focus();
			return false;
		}
		else if(document.wuliuform.S_year.value>document.wuliuform.E_year.value )//&& document.wuliuform.S_moon.value>document.wuliuform.E_moon.value
		{alert("起始月份不能超过终止月份！");
		return false;
			}else{
			
	return true;}
}
</script>
<form id="wuliuform" name="wuliuform" method="post" action="" onSubmit="return checkform();">
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
              <option value="01" <%if S_moon=1 then %>selected="selected"<%end if%>>1</option>
              <option value="02" <%if S_moon=2 then %>selected="selected"<%end if%>>2</option>
              <option value="03" <%if S_moon=3 then %>selected="selected"<%end if%>>3</option>
              <option value="04" <%if S_moon=4 then %>selected="selected"<%end if%>>4</option>
              <option value="05" <%if S_moon=5 then %>selected="selected"<%end if%>>5</option>
              <option value="06" <%if S_moon=6 then %>selected="selected"<%end if%>>6</option>
              <option value="07" <%if S_moon=7 then %>selected="selected"<%end if%>>7</option>
              <option value="08" <%if S_moon=8 then %>selected="selected"<%end if%>>8</option>
              <option value="09" <%if S_moon=9 then %>selected="selected"<%end if%>>9</option>
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
              <option value="01" <%if E_moon=1 then %>selected="selected"<%end if%>>1</option>
              <option value="02" <%if E_moon=2 then %>selected="selected"<%end if%>>2</option>
              <option value="03" <%if E_moon=3 then %>selected="selected"<%end if%>>3</option>
              <option value="04" <%if E_moon=4 then %>selected="selected"<%end if%>>4</option>
              <option value="05" <%if E_moon=5 then %>selected="selected"<%end if%>>5</option>
              <option value="06" <%if E_moon=6 then %>selected="selected"<%end if%>>6</option>
              <option value="07" <%if E_moon=7 then %>selected="selected"<%end if%>>7</option>
              <option value="08" <%if E_moon=8 then %>selected="selected"<%end if%>>8</option>
              <option value="09" <%if E_moon=9 then %>selected="selected"<%end if%>>9</option>
              <option value="10" <%if E_moon=10 then %>selected="selected"<%end if%>>10</option>
              <option value="11" <%if E_moon=11 then %>selected="selected"<%end if%>>11</option>
              <option value="12" <%if E_moon=12 then %>selected="selected"<%end if%>>12</option>
            </select>
            <%        set srs=server.createobject("adodb.recordset")
        ssql="select * from Supervisor order by Sup_Orderid desc,Sup_id desc"

        srs.open ssql,conn,3,3%>
            <select name="Sup_id" id="Sup_id">
            <option value="0" >项目主管</option>
            <%do while not srs.eof %>
                <option value="<%=srs("Sup_id")%>" <%if Sup_id=srs("Sup_id") then %>selected="selected"<%end if%>><%=srs("Sup_name")%></option>
                <% srs.movenext
 loop%>
            </select>
            <%	 srs.close
	 set srs=nothing%>
            <input type="hidden" name="skey" id="skey" value="1">
            <input type="submit" name="button" id="button" value="搜索" />　<a href="?<%=urlload%>&Riframe=4&time=Clear">查看全部</a></form>