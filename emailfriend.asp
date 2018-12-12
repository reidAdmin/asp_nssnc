<!--#include file="vsadmin/db_conn_open.asp"-->
<!--#include file="vsadmin/inc/languagefile.asp"-->
<!--#include file="vsadmin/includes.asp"-->
<%
savecodepage=response.codepage
if inlinepopups=TRUE then
	if lcase(adminencoding)<>"utf-8" then response.codepage=65001
	response.charset="utf-8"
end if
%>
<!--#include file="vsadmin/inc/incfunctions.asp"-->
<!--#include file="vsadmin/inc/incemail.asp"-->
<%	if inlinepopups<>TRUE then %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title><%=xxEmFrnd%></title>
<link rel="stylesheet" type="text/css" href="style.css" />
<meta name="robots" content="noindex,nofollow" />
</head>
<body style="margin: 5px 5px 5px 5px;">
<%	else
Response.Buffer = True
Response.Expires = 60
Response.Expiresabsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<table width="500" border="0" id="emftable" style="position:absolute">
	<tr>
	  <td align="center" id="efrcell">
<%
	end if
if multiemfblockmessage="" then multiemfblockmessage="I'm sorry. We are experiencing temporary difficulties at the moment. Please try again later."
if request("askq")="1" AND useaskaquestion=TRUE then isaskquestion=TRUE else isaskquestion=FALSE
function checkemfuserblock()
	if blockmultiemf="" then blockmultiemf=20
	multiemfblocked=FALSE
	theip = trim(replace(left(request.servervariables("REMOTE_ADDR"), 48), "'", ""))
	if theip = "" then theip = "none"
	if blockmultiemf<>"" then
		cnn.Execute("DELETE FROM multibuyblock WHERE lastaccess<" & datedelim & VSUSDateTime(Now()-1) & datedelim)
		sSQL = "SELECT ssdenyid,sstimesaccess FROM multibuyblock WHERE ssdenyip = '" & "EMF " & theip & "'"
		rs.Open sSQL,cnn,0,1
		if NOT rs.EOF then
			cnn.Execute("UPDATE multibuyblock SET sstimesaccess=sstimesaccess+1,lastaccess=" & datedelim & VSUSDateTime(Now()) & datedelim & " WHERE ssdenyid=" & rs("ssdenyid"))
			if rs("sstimesaccess") >= blockmultiemf then multiemfblocked=TRUE
		else
			cnn.Execute("INSERT INTO multibuyblock (ssdenyip,lastaccess) VALUES ('" & "EMF " & theip & "'," & datedelim & VSUSDateTime(Now()) & datedelim & ")")
		end if
		rs.Close
	end if
	if theip = "none" then
		sSQL = "SELECT "&IIfVr(mysqlserver<>true,"TOP 1","")&" dcid FROM ipblocking"&IIfVr(mysqlserver=true," LIMIT 0,1","")
	else
		sSQL = "SELECT dcid FROM ipblocking WHERE (dcip1=" & ip2long(theip) & " AND dcip2=0) OR (dcip1 <= " & ip2long(theip) & " AND " & ip2long(theip) & " <= dcip2 AND dcip2 <> 0)"
	end if
	rs.Open sSQL,cnn,0,1
	if NOT rs.EOF then multiemfblocked = TRUE
	rs.Close
	checkemfuserblock = multiemfblocked
end function
	if request.form("posted")="1" then
		success=TRUE
		referer = request.servervariables("HTTP_REFERER")
		host = request.servervariables("HTTP_HOST")
		if instr(referer, host)=0 then
			xxEFThk="<strong><font color=""#FF0000"">I'm sorry but your email could not be sent at this time.</font></strong>"
		else
			if htmlemails=true then emlNl = "<br />" else emlNl=vbCrLf
			theprodid = trim(left(request.form("efid"),50))
			Set rs = Server.CreateObject("ADODB.RecordSet")
			Set cnn=Server.CreateObject("ADODB.Connection")
			cnn.open sDSN
			if useemailfriend<>TRUE AND useaskaquestion<>TRUE then
				xxEFThk="<strong><font color=""#FF0000"">Email Friend / Ask a Question not enabled.</font></strong>"
			elseif checkemfuserblock() then
				xxEFThk="<strong><font color=""#FF0000"">" & multiemfblockmessage & "</font></strong>"
				response.status = "403 Forbidden"
				response.end
			else
				sSQL="SELECT adminEmail,smtpserver,emailUser,emailPass,adminStoreURL,emailObject FROM admin WHERE adminID=1"
				rs.Open sSQL,cnn,0,1
				emailAddr = rs("adminEmail")
				themailhost = Trim(rs("smtpserver")&"")
				theuser = Trim(rs("emailUser")&"")
				thepass = Trim(rs("emailPass")&"")
				adminStoreURL = rs("adminStoreURL")
				if (left(LCase(adminStoreURL),7) <> "http://") AND (left(LCase(adminStoreURL),8) <> "https://") then
					adminStoreURL = "http://" & adminStoreURL
				end if
				if Right(adminStoreURL,1) <> "/" then adminStoreURL = adminStoreURL & "/"
				emailObject = rs("emailObject")
				rs.Close
				if isaskquestion AND useaskaquestion=TRUE then
					friendsemail = emailAddr
				elseif useemailfriend=TRUE AND len(request.form("friendsemail"))<50 then
					friendsemail = left(request.form("friendsemail"),50)
				else
					friendsemail = ""
				end if
				yourname = left(request.form("yourname"),50)
				youremail = left(request.form("youremail"),50)
				yourcomments = replace(trim(left(request.form("yourcomments"),2000)),vbCrLf,emlNl)
				if isaskquestion then
					seBody = xxAskQue & ": " & yourname & emlNl & emlNl & yourcomments & emlNl
					thesubject=xxAsqSub
				else
					seBody = xxEFYF1 & yourname & " (" & youremail & ")" & xxEFYF2
					if trim(request.form("yourcomments"))<>"" then
						seBody = seBody & xxEFYF3 & emlNl
						seBody = seBody & yourcomments & emlNl
					else
						seBody = seBody & "." & emlNl
					end if
					produrl=""
					if theprodid<>"" then
						sSQL = "SELECT pID,"&getlangid("pName",1)&",pStaticPage,pStaticURL FROM products WHERE pID='" & escape_string(theprodid) & "'"
						rs.Open sSQL,cnn,0,1
						if lcase(adminencoding)<>"utf-8" then response.codepage=savecodepage
						if NOT rs.EOF then produrl=getdetailsurl(rs("pID"),rs("pStaticPage"),rs(getlangid("pName",1)),trim(rs("pStaticURL")&""),"","")
						if lcase(adminencoding)<>"utf-8" then response.codepage=65001
						rs.Close
					end if
					if htmlemails=TRUE then
						storeLink = adminStoreURL
						if Trim(Request.Form("efid")) <> "" then storeLink = storeLink & produrl
						seBody = seBody & emlNl & "<a href=""" & storeLink & """>" & storeLink & "</a>"
					else
						seBody = seBody & emlNl & adminStoreURL
						if Trim(Request.Form("efid")) <> "" then seBody = seBody & produrl
					end if
					thesubject=yourname & xxEFRec
				end if
				seBody = seBody & emlNl
				if friendsemail<>"" then call DoSendEmailEO(friendsemail,emailAddr,youremail,thesubject,seBody,emailObject,themailhost,theuser,thepass)
			end if
			cnn.Close
			set rs = nothing
			set cnn = nothing
		end if
%>
<br />
  <table class="cobtbl emftbl" border="0" cellspacing="1" cellpadding="3" width="<%=IIfVr(inlinepopups=TRUE,"500","100%")%>">
	<tr>
	  <td class="cobll emfll"colspan="2" align="center" width="100%"><p>&nbsp;</p>
	  <p><%=IIfVr(isaskquestion,xxAsqThk,xxEFThk)%></p>
	  <p><%=xxClkClo%></p>
	  <p>&nbsp;</p>
	  <%=imageorbutton(imgefclose, xxClsWin, "", IIfVr(inlinepopups=TRUE,"document.body.removeChild(document.getElementById('efrdiv'))","javascript:self.close()"), TRUE)%>
	  <p>&nbsp;</p>
	  </td>
	</tr>
  </table>
<%	else
	if inlinepopups<>TRUE then call emailfriendjavascript() %>
<form id="efform" method="post" action="emailfriend.asp" onsubmit="return efformvalidator(this)">
  <input type="hidden" name="posted" value="1" />
  <input type="hidden" id="efid" name="efid" value="<%=server.htmlencode(Request.QueryString("id"))%>" />
  <input type="hidden" id="askq" name="askq" value="<%=IIfVr(isaskquestion,"1","")%>" />
  <table class="cobtbl emftbl" border="0" cellspacing="1" cellpadding="7" width="<%=IIfVr(inlinepopups=TRUE,"500","100%")%>">
	<tr>
	  <td class="cobhl emfhl" align="center" width="100%" height="30"><%=IIfVr(isaskquestion,xxAskQue,xxEmFrnd)%></td>
	</tr>
	<tr>
		<td class="cobll emfll" width="100%" align="left"><%=IIfVr(isaskquestion,xxAQBlr,xxEFBlr)%><br />
		<br /><font color="#FF0000">*</font><%=xxEFNam%><br /><input type="text" id="yourname" name="yourname" size="30" /><br />
		<font color="#FF0000">*</font><%=xxEFEm%><br /><input type="text" id="youremail" name="youremail" size="30" /><br />
<%		if NOT isaskquestion then %> 
		<font color="#FF0000">*</font><%=xxEFFEm%><br /><input type="text" id="friendsemail" name="friendsemail" size="30" /><br />
<%		else
			theproduct=trim(left(request.querystring("id"),50))
			Set rs = Server.CreateObject("ADODB.RecordSet")
			Set cnn=Server.CreateObject("ADODB.Connection")
			cnn.open sDSN
			sSQL = "SELECT pName FROM products WHERE pID='" & escape_string(theproduct) & "'"
			rs.Open sSQL,cnn,0,1
			if NOT rs.EOF then
				theproduct = theproduct & " - " & rs("pName")
			end if
			rs.Close
			cnn.Close
			set rs = nothing
			set cnn = nothing
		end if %>
		<font color="#FF0000">*</font><%=xxEFCmt%><br /><textarea id="yourcomments" name="yourcomments" cols="46" rows="6"><%=IIfVr(isaskquestion,htmlspecials(replace(xxAskCom,"%nl%",vbCrLf) & theproduct),"")%></textarea>
		<p align="center"><%
		if inlinepopups=TRUE then
			print imageorbutton(imgefsend, xxSend, "", "dosendefdata()", TRUE)
		else
			print imageorsubmit(imgefsend, xxSend, "")
		end if
		print "&nbsp;&nbsp;"
		print imageorbutton(imgefclose, xxClsWin, "", IIfVr(inlinepopups=TRUE,"document.body.removeChild(document.getElementById('efrdiv'))","javascript:self.close()"), TRUE)
%></p>
      </td>
	</tr>
  </table>
</form>
<%		if inlinepopups=TRUE then %>
	</td>
  </tr>
</table>
<%		end if
	end if
	if inlinepopups<>TRUE then %>
</body>
</html>
<%	end if %>