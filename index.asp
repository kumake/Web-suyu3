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
			<!--#include file="head.asp" --></td>
	</tr>
	<tr>
		<td  width="961" height="667">
		<table id="__01" width="961" height="667" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td>
					<!--#include file="left.asp" -->
					</td>
					<td>
					<table id="__01" width="716" height="667" border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td>
							<table id="__01" width="716" height="181" border="0" cellpadding="0" cellspacing="0">
								<tr>
									<td>
										<img src="images/index_qq_01.jpg" width="53" height="181" alt=""></td>
									<td background="images/index_qq_02.jpg" width="653" height="181">
									<!--内容开始 -->
											<script src="JS/MSClass.js"></script>
									<div id="marqueediv6" style=" text-align:center;width:640px;height:170px;margin-left:13px; padding-top:5px;">
								  <table width="100%" border="0" cellspacing="0" cellpadding="0">
									  <tr>
									  
									   <%
			
							set rs = UICon.QUery("select top 20 * from user_case order by hots desc ")
							if rs.recordcount<>0 then
							do while not rs.eof
							'''''''''怎么分多列''''''
							''该页面采用DIV+css。所以分列实现起来非常方便不需要改页面
							''只需要改变css中 procontent 标签下的li的宽度即可
							''一列 只要procontent 的宽度和li的宽度一致
							''需要几列 只要将li的宽度设置为 procontent的几分之几
						   %>
										<td width="122"><table width="122" border="0" align="center" cellpadding="0" cellspacing="0"  height="122">
											<tr>
											  <td width="122"><a href="case_in.asp?categoryID=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>"><img src="<%=rs("Field_picture")%>"  height="130" ; width="150"   border="0" style="margin-top:5px"/></a>
											  <a href="case__in.asp?categoryID=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" style="display:block; text-align:center; line-height:20px; color:#000; margin-top:5px;"><%=rs("title")%></a>								  </td>
											</tr>
										</table></td>
										<td width="10">&nbsp;</td>
						  <%
							rs.movenext
							loop
							end if
							%>              
									  </tr>
					  </table>
					</div>
									<script>new Marquee("marqueediv6",2,1,640,170,20,0,0)</script>
									<!--内容结束 -->
									</td>
									<td>
										<img src="images/index_qq_03.jpg" width="10" height="181" alt=""></td>
								</tr>
							</table>
							</td>
						</tr>
						<tr>
							<td>
							<table id="__01" width="716" height="213" border="0" cellpadding="0" cellspacing="0">
								<tr>
									<td>
										<table id="__01" width="358" height="213" border="0" cellpadding="0" cellspacing="0">
											<tr>
												<td>
													<img src="images/index_ss_01.jpg" width="358" height="36" alt=""></td>
											</tr>
											<tr>
												<td width="358" height="177">
												<ul id="index_news">
													<%
													set rs = UICon.Query("select top 6 * from user_news order by id desc")
													do while not rs.eof
												
												%>
													<li>　　　　<a  href="news_in.asp?categoryid=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" title="<%=rs("title")%>"  target="_blank" ><%
													if len(rs("title"))>14 then	
													response.write left(rs("title"),13)&"..."
													else
													response.write rs("title")
													end if
													%></a></li>	
												<%
													rs.movenext
													loop
													rs.close
													set rs=nothing
												%>
												</ul>
												</td>
											</tr>
										</table>
									</td>
									<td>
									<table id="__01" width="358" height="213" border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td>
												<img src="images/index_hh_01.jpg" width="358" height="37" alt=""></td>
										</tr>
										<tr>
											<td width="358" height="176">
											<div id="index_gsjs">
											<p>我公司自 1997 年 3 月创办至今，规模不断扩大，现有员工 185 人，其中专业维修技术人员 171 人，专业企划人员 9 人。其中高级职称的技术人员占总人数的 63% ，是一个具有高素质，高质量的服务体系，能确保客户服务的及时性和经济性，公司还具有高科技信息网络，在维修接单及售后服务过程中全程电脑化管理，因此业务不断扩大，行业声望也日益提高，并先后与媒介行业</p>
											</div>
											</td>
										</tr>
									</table>
									</td>
								</tr>
							</table>
							</td>
						</tr>
						<tr>
							<td>
								<img src="images/index_dd_03.jpg" width="716" height="36" alt=""></td>
						</tr>
						<tr>
							<td width="716" height="237">

								<%
								set rsn = UICon.Query("select top 2 * from Sp_dictionary where  categoryID = 12 and parentID=0 order by  IndexID") 
								do while not rsn.eof	
								%>
								<div class="index_fwfw">
						<div class="fwfw_head"><%=rsn("categoryname")%></div>
									<!--开始 -->
										<%
										dim cateid:cateid=rsn("id") 
										dim ssi:ssi=0
										set rs = UICon.Query("select top 9 * from user_pro where isdel=0 and categoryID="&cateid&" order by id desc")
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
									</div>									
								<%
									rsn.movenext
									loop
									rsn.close
									set rsn=nothing
								%>
							</td>
						</tr>
					</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
			<!--#include file="footer.asp" --></td>
	</tr>
</table>
</div>
</body>
</html>
