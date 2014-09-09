<!--#include file="conn.asp"-->
<!--#include file="sp_inc/class_page.asp"-->
<!--#include file="plugIn/Setting.Config.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=config_sitename%></title>
<meta name="keywords" content="<%=config_seokeyword%>">
<meta name="description" content="<%=config_seocontent%>">
<link href="css/public.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.proLi{ width:160px; line-height:30px; border-bottom:#CCCCCC solid 1px; display:block; background:url(images/li.jpg) no-repeat 30px 50%; padding-left:50px; margin-left:32px;}
 -->
</style>
</head>
<body>
<div id="container">
<table id="__01" width="961" height="1157" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td>
		<!--#include file="head.asp" -->
		</td>
	</tr>
	<tr>
		<td>
		<table id="__01" width="961" height="667" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td>
		<!--#include file="left.asp" -->
		</td>
		<td>
		<table id="__01" width="716" height="667" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td background="images/neiye_dd_01.jpg" width="716" height="46">
				<div style="padding-top:22px; margin-left:30px; font-size:14px; font-weight:bold;color:#fff;">服务范围</div>				</td>
			</tr>
			<tr>
				<td  width="716" height="621" >
				<div id="neiye_main">
					<div id="neiye_text">
								<%
								set rsn = UICon.Query("select  * from Sp_dictionary where  categoryID = 12 and parentID=0 order by  IndexID") 
								do while not rsn.eof	
								%>
						<div class="fwfw_head"><%=rsn("categoryname")%></div>
									<!--开始 -->
										<%
										dim cateid:cateid=rsn("id") 
										dim ssi:ssi=0
										set rs = UICon.Query("select  * from user_pro where isdel=0 and categoryID="&cateid&" order by id desc")
										do while not rs.eof	
										ssi=ssi+1		
										%>
						<p><%=ssi%>、<%=rs("title")%></p>	
										<%
											rs.movenext
											loop
											rs.close
											set rs=nothing
										%>
									<!--结束 -->										
								<%
									rsn.movenext
									loop
									rsn.close
									set rsn=nothing
								%>								
					</div>
				</div>				</td>
			</tr>
		</table>
		</td>
	</tr>
</table>
		
		</td>
	</tr>
	<tr>
		<td>
		<!--#include file="footer.asp" -->
		</td>
	</tr>
</table>
</div>
</body>
</html>
