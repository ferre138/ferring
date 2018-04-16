<!--#include file="ewcfg9.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<%

' Get Menu Text
Function GetMenuText(Id, Text)
	GetMenuText = Language.MenuPhrase(Id, "MenuText")
	If GetMenuText = "" Then GetMenuText = Text
End Function
%>
<!-- Begin Main Menu -->
<div class="aspmaker">
<%

' Initialize language object
Set Language = New cLanguage
Call Language.LoadPhrases()

' Generate all menu items
Dim RootMenu
Set RootMenu = new cMenu
RootMenu.Id = "RootMenu"
RootMenu.IsRoot = True
RootMenu.AddMenuItem 3, GetMenuText("3", "Orders"), "Orderslist.asp", -1, "", "", IsLoggedIn(), False
RootMenu.AddMenuItem 1, GetMenuText("1", "Customers"), "Customerslist.asp", -1, "", "", IsLoggedIn(), False
RootMenu.AddMenuItem 12, GetMenuText("12", "Store Settings"), "", -1, "", "", IsLoggedIn(), True
RootMenu.AddMenuItem 2, GetMenuText("2", "Logins"), "Loginslist.asp", 12, "", "", IsLoggedIn(), False
RootMenu.AddMenuItem 5, GetMenuText("5", "Products"), "Productslist.asp", 12, "", "", IsLoggedIn(), False
RootMenu.AddMenuItem 6, GetMenuText("6", "Province"), "Provincelist.asp", 12, "", "", IsLoggedIn(), False
RootMenu.AddMenuItem 28, GetMenuText("28", "Orders Report"), "reportDaily.asp", -1, "", "", True, False
RootMenu.AddMenuItem &HFFFFFFFE, Language.Phrase("ChangePwd"), "changepwd.asp", -1, "", "", (IsLoggedIn() And Not IsSysAdmin()), False
RootMenu.AddMenuItem &HFFFFFFFF, Language.Phrase("Logout"), "logout.asp", -1, "", "", IsLoggedIn(), False
RootMenu.AddMenuItem &HFFFFFFFF, Language.Phrase("Login"), "login.asp", -1, "", "", (Not IsLoggedIn() And Right(Request.ServerVariables("URL"), Len("login.asp")) <> "login.asp"), False
Call RootMenu.Render(False)
Set RootMenu = Nothing
%>
</div>
<!-- End Main Menu -->
<script type="text/javascript">
<!--
// init the menu
var RootMenu = new YAHOO.widget.Menu("RootMenu", { position: "static", hidedelay: 750, lazyload: true });
RootMenu.render();        
//-->
</script>
