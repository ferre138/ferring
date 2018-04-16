<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<title><%= Language.ProjectPhrase("BodyTitle") %></title>
<link rel="stylesheet" type="text/css" href="<%= ew_YuiHost %>build/menu/assets/skins/sam/menu.css">
<link rel="stylesheet" type="text/css" href="css/ewmenu.css">
<link rel="stylesheet" type="text/css" href="<%= ew_YuiHost %>build/tabview/assets/skins/sam/tabview.css">
<link rel="stylesheet" type="text/css" href="<%= ew_YuiHost %>build/container/assets/skins/sam/container.css">
<link rel="stylesheet" type="text/css" href="<%= ew_YuiHost %>build/resize/assets/skins/sam/resize.css">
<link rel="stylesheet" type="text/css" href="<%= EW_PROJECT_STYLESHEET_FILENAME %>">
<script type="text/javascript" src="<%= ew_YuiHost %>build/utilities/utilities.js"></script>
<script type="text/javascript" src="<%= ew_YuiHost %>build/tabview/tabview-min.js"></script>
<script type="text/javascript" src="<%= ew_YuiHost %>build/container/container-min.js"></script>
<script type="text/javascript" src="<%= ew_YuiHost %>build/resize/resize-min.js"></script>
<script type="text/javascript" src="<%= ew_YuiHost %>build/menu/menu.js"></script>
<script type="text/javascript">
<!--
var EW_LANGUAGE_ID = "<%= gsLanguage %>";
var EW_DATE_SEPARATOR = "/"; 
if (EW_DATE_SEPARATOR == "") EW_DATE_SEPARATOR = "/"; // Default date separator
var EW_UPLOAD_ALLOWED_FILE_EXT = "gif,jpg,jpeg,bmp,png,doc,xls,pdf,zip"; // Allowed upload file extension
var EW_FIELD_SEP = ", "; // Default field separator
// Ajax settings
var EW_RECORD_DELIMITER = "\r";
var EW_FIELD_DELIMITER = "|";
var EW_LOOKUP_FILE_NAME = "ewlookup9.asp"; // Lookup file name
var EW_AUTO_SUGGEST_MAX_ENTRIES = <%= EW_AUTO_SUGGEST_MAX_ENTRIES %>; // Auto-Suggest max entries
// Common JavaScript messages
var EW_ADDOPT_BUTTON_SUBMIT_TEXT = "<%= ew_JsEncode2(ew_BtnCaption(Language.Phrase("AddBtn"))) %>";
var EW_EMAIL_EXPORT_BUTTON_SUBMIT_TEXT = "<%= ew_JsEncode2(ew_BtnCaption(Language.Phrase("SendEmailBtn"))) %>";
var EW_BUTTON_CANCEL_TEXT = "<%= ew_JsEncode2(ew_BtnCaption(Language.Phrase("CancelBtn"))) %>";
var ewTooltipDiv;
var ew_TooltipTimer = null;
//-->
</script>
<script type="text/javascript" src="js/ew9.js"></script>
<script type="text/javascript" src="js/ewvalidator.js"></script>
<script type="text/javascript" src="js/userfn8.js"></script>
<script type="text/javascript">
<!--
<%= Language.ToJSON() %>
//-->
</script>
<script language="JavaScript" type="text/javascript">
<!--
// Write your client script here, no need to add script tags.
//-->
</script>
<link rel="shortcut icon" type="image/png" href="<%= ew_ConvertFullUrl("asp80.png") %>"><link rel="icon" type="image/png" href="<%= ew_ConvertFullUrl("asp80.png") %>">
<meta name="generator" content="ASPMaker v9.0.2">
</head>
<body class="yui-skin-sam">
<div class="ewLayout">
	<!-- header (begin) --><!-- *** Note: Only licensed users are allowed to change the logo *** -->
  <div class="ewHeaderRow"><img src="images/vslBanner.png" alt="" border="0"></div>
	<!-- header (end) -->
	<!-- content (begin) -->
  <table cellspacing="0" class="ewContentTable">
		<tr>
			<td class="ewMenuColumn">
			<!-- left column (begin) -->
<% Server.Execute("ewmenu.asp") %>
			<!-- left column (end) -->
			</td>
		<td class="ewContentColumn">
			<!-- right column (begin) -->
				<p class="aspmaker ewTitle"><b><%= Language.ProjectPhrase("BodyTitle") %></b></p>
