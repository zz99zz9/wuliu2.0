<%

Exh_code=request.QueryString("Exh_code")
ECount=request.QueryString("ECount")
ECount2=request.QueryString("ECount2")
Exh_id=request.QueryString("Exh_id")
Riframe=request.QueryString("Riframe")
zhlb=request.QueryString("zhlb")
timeClear=request.QueryString("time")
response.cookies("keyword")=request("keyword")
keyword=request.cookies("keyword")

response.cookies("S_year")=request("S_year")
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
E_moon=request.cookies("E_moon")
if Riframe="" then Riframe=0
if zhlb="" then zhlb=0
if ECount="" then ECount=0
if ECount2="" then ECount2=0
urlload="Exh_code="&Exh_code&"&ECount="&ECount&"&ECount2="&ECount2&"&Exh_id="&Exh_id&"&zhlb="&zhlb'&"&Riframe="&Riframe
'Exh_codeչ���� ��ECount��ǰչ����  ��ECount2��ǰչ����2  �� Exh_idչ��id ��zhlb�б�
%>
<!--#include virtual="head.asp"-->
<%if request.cookies("wuliuv")=1 then
response.Redirect("index2.asp")
end if%>
<script language="JavaScript" type="text/javascript"> 
function ChangeDiv(divId,divName,divName2,zDivCount) 
{ 
for(i=0;i<=zDivCount;i++) 
{ 
document.getElementById(divName+i).style.display="none"; 
document.getElementById(divName2+i).className="l"; 
//�����еĲ㶼���� 
} 
document.getElementById(divName+divId).style.display="block"; 
document.getElementById(divName2+divId).className="l1"; 
//��ʾ��ǰ�� 
} 
function ShowDiv(divName) 
{ 

document.getElementById(divName).style.display="block"; 

} 
function HiddenDiv(divName) 
{ 

document.getElementById(divName).style.display="none"; 

} 
</script> 
<div id="Mainbottom">
<div id="Mainleftmenu">
<div class="Topmenu">
<ul>
<li class="l1" onClick="javascript:ChangeDiv('0','zhlb','zhbt',1)" id="zhbt0">��������</li>
<li class="l" onClick="javascript:ChangeDiv('1','zhlb','zhbt',1)" id="zhbt1">չ���б�</li>
<li class="l" id="zhbt2" style="display:none;"></li>
</ul>
</div>
<div class="clear"></div>
<!--��������-->
<div id="zhlb0" class="zhlb"><br /><!--#include virtual="inc/left_sstj.asp"--><div class="clear"></div></div>
<!--չ���б�-->
<div id="zhlb1" class="zhlb"><br /><!--#include virtual="inc/left_zhlb.asp"--><div class="clear"></div></div>
<!--չ���б�-->
<div id="zhlb2" class="zhlb" style="display:none;"><br /><div class="clear"></div></div>

</div>
    <div id="MainrightContent">
    <%if Exh_id="" then%><div class="zhlbrhidden"><br /><br /><br /><br />
   
   ���� �������ұ��Ӵ�ѡ��һ��չ���"<b>ҵ�����</b>"
    </div>
    <%end if%>
          <div class="Topmenu">
        <ul>
          <li class="l1" onClick="javascript:ChangeDiv('0','zhlbr','zhb',3)" id="zhb0">�������</li>
          <li class="l" onClick="javascript:ChangeDiv('1','zhlbr','zhb',3)" id="zhb1">֧������</li>
          <li class="l" onClick="javascript:ChangeDiv('2','zhlbr','zhb',3)" id="zhb2">��������</li>
          <li class="l" onClick="javascript:ChangeDiv('3','zhlbr','zhb',3)" id="zhb3">Ӧ�տ��</li>
       <!--   <li class="l" onClick="javascript:ChangeDiv('4','zhlbr','zhb',4)" id="zhb4">ҵ��һ����</li>-->
        </ul>
      </div>
      <div class="clear"></div>
      <!--�������-->
      <div id="zhlbr0" class="zhlbr">
<!--#include virtual="inc/Revenue.asp"-->
      </div>
      <!--֧������-->
            <div id="zhlbr1" class="zhlbr">
<!--#include virtual="inc/Expense.asp"-->
      </div>
      <!--��������-->
            <div id="zhlbr2" class="zhlbr">
<!--#include virtual="inc/Settlement3.asp"-->
      </div>
      <!--Ӧ�տ��-->
      <div id="zhlbr3" class="zhlbr">
<!--#include virtual="inc/Accounts.asp"-->
      </div>
      <!--ҵ��һ����
      <div id="zhlbr4" class="zhlbr">-->
<!-- # include virtual="inc/Chart.asp"-->
      <!--- </div>
     -------->
 

    </div>
</div>
<script type="text/javascript" src="js/jquery.js"></script>

<script type="text/javascript" src="js/search.js"></script>
</body>
</html>
