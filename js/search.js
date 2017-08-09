var searchReq=createAjaxObj();
function createAjaxObj()
{
	var httprequest=false;
	if(window.XMLHttpRequest)
	{
		httprequest=new XMLHttpRequest();
		if(httprequest.overrideMimeType)
			httprequest.overrideMimeType('text/xml');
	}
	else if (window.ActiveXObject)
	{
		//IE
		try
		{
			httprequest=new ActiveXObject("Msxml2.XMLHTTP");
		}
		catch (e)
		{
			try
			{
				httprequest=new ActiveXObject("Microsoft.XMLHTTP");
			}
			catch (e)
			{
			}
		}
	}
	return httprequest
}

function searchSuggest()
{
	if ($('#Rev_customer').val())
	{
		var str=escape($('#Rev_customer').val());
		url="js/search.asp?search="+str+ "&t=" +  new Date().getTime();
		searchReq.open("get",url);
		searchReq.onreadystatechange=handleSearchSuggest;
		searchReq.send(null);	
	}
	else
	{
		$(this).next(".search_suggest").html("");
		$(this).next(".search_suggest").css("display","none");
	}
	
	
}
function searchSuggest3()
{
	if ($('#Rev_customer2').val())
	{
		var str=escape($('#Rev_customer2').val());
		url="js/search.asp?search="+str+ "&t=" +  new Date().getTime();
		searchReq.open("get",url);
		searchReq.onreadystatechange=handleSearchSuggest;
		searchReq.send(null);	
	}
	else
	{
		$(this).next(".search_suggest").html("");
		$(this).next(".search_suggest").css("display","none");
	}
	
	
}
function searchSuggest2()
{
	if ($('#Exp_customer').val())
	{
		var str=escape($('#Exp_customer').val());
		url="js/search.asp?search="+str+ "&t=" +  new Date().getTime();
		searchReq.open("get",url);
		searchReq.onreadystatechange=handleSearchSuggest;
		searchReq.send(null);	
	}
	else
	{
		$(this).next(".search_suggest").html("");
		$(this).next(".search_suggest").css("display","none");
	}
	
	
}
function searchSuggest4()
{
	if ($('#Exp_customer2').val())
	{
		var str=escape($('#Exp_customer2').val());
		url="js/search.asp?search="+str+ "&t=" +  new Date().getTime();
		searchReq.open("get",url);
		searchReq.onreadystatechange=handleSearchSuggest;
		searchReq.send(null);	
	}
	else
	{
		$(this).next(".search_suggest").html("");
		$(this).next(".search_suggest").css("display","none");
	}
	
	
}
//download by http://www.codefans.net
function handleSearchSuggest()
{
	if(searchReq.readyState==4)
	{		var ss="";
			$(".search_suggest").html("");
			s0=searchReq.responseText.length;		
			if (s0>0)
			{
				xmldoc=searchReq.responseXML;	
				var message_nodes=xmldoc.getElementsByTagName("message");
				var n_messages=message_nodes.length;				
				if (n_messages<=0)
				{
					$(".search_suggest").html("");
					$(".search_suggest").css("display","block");
				}
			    else
				{ 
					$(".search_suggest").css("display","block");
					for (i=0;i<n_messages;i++ )
					{var suggg=""
						var suggest='<div onmouseover="javascript:suggestOver(this);"';	
						suggest+='onmouseout="javascript:sugggestOut(this);"';
						suggest+='onclick="javascript:setSearch(this.innerHTML);"';
						suggest +='class="suggest_link">'+message_nodes[i].getElementsByTagName("text")[0].firstChild.data+'</div>';
						ss +=suggest;
						$(".search_suggest").html(ss);	
					}	
								
				}
			}
			else
			{
				$(".search_suggest").html("");
				$(".search_suggest").css("display","none");
			}		
	}
	else
	{
		//alert('��������ʧ��');
	}
}

function suggestOver(div_value)
{
	div_value.className='suggest_link_over';
}

function sugggestOut(div_value)
{
  div_value.className='suggest_link';
}

function setSearch(div_value)
{
   $(".tete").val(div_value);
   $(".search_suggest").html("");
   $(".search_suggest").css("display","none");
}