<link rel="stylesheet" type="text/css" media="all" href="styles.css" >
<%
path= request.querystring("path")
title=request.querystring("title")

%>
<style type="text/css">
body {margin:0px;}
</style> 
<body>
<form action="mailprocess.asp" method="post" name="myForm" id="myForm">
<table width="673" height="395" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="50" align="left" valign="bottom" class="bodycopy"> <div align="center"><font size="3"><strong>I would like to send the current animation friend(s)</strong></font></div></td>
    <td width="26" rowspan="3" align="left" valign="top">&nbsp;</td>
  </tr>
  <tr>
    <td width="621" align="left" valign="middle"><div align="center">
      <table border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td colspan="2"><table border="0" cellspacing="0" cellpadding="0">
              <tr class="bodycopy">
                <td width="121" height="30"><strong><strong>Your Name</strong></strong></td>
                <td height="30"><label>
                  <input name="FullName" type="text" id="FullName" size="35" maxlength="35" />
                </label></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td colspan="2"><table border="0" cellspacing="0" cellpadding="0">
              <tr class="bodycopy">
                <td width="121" height="30"><strong>Your Email</strong></td>
                <td height="30"><label>
                  <input name="Email" type="text" id="Email" size="35" maxlength="35" />
                </label></td>
              </tr>
            </table></td>
          </tr>
    
      <tr>
            <td colspan="2"><table border="0" cellspacing="0" cellpadding="0">
              <tr class="bodycopy">
                <td width="121" height="30"><strong>Friend's Email</strong><font size="-1"><br />
                  (to add multiple friends, simply separate addresses with a comma)
                  </font>              <table cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="180" valign="top"> </td>
                    </tr>
                  </table></td>
                <td height="30" valign="top"><label>
                  <textarea name="FriendsEmail" cols="50" rows="4" id="FriendsEmail"></textarea>
                </label></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td  align="center" valign="middle">&nbsp;</td>
            <td align="left" valign="middle"><div align="center">
              <label>
                <input type="submit" name="send" id="send" value="Submit" />
                <input name="hSrcForm" type="hidden" id="hSrcForm" value="EmailSWF" />
		          <input name="hPath" type="hidden" id="hPath" value="<%=path%>" />
                  <input name="hPageTitle" type="hidden" id="hPageTitle" value="<%=sTitle%>" />
              </label>
            </div></td>
          </tr>
      </table>
    </div></td>
  </tr>
  <tr>
    <td height=30 colspan="3" bgcolor="#3c74b6" ><div align="center"><strong><font color="#FFFFFF">Email movie : <%=title%></font></strong> </div></td>
  </tr>
</table>
</form>
</body>
