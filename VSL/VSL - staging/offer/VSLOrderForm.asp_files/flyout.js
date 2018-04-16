// DHTML Menu for Graphical Template
// (C)2005 e.World Technology Ltd.
// 2005/09/15 Added support for FF1+/NS6+

var global = window.document
var MENU_BORDER_COLOR = '#999999'
var MENU_CURRENTPAGE_COLOR = '#ffffff'
var MENU_MOUSEOVER_COLOR = '#cccccc'
var MENU_MOUSEDOWN_COLOR = '#999999'

function normalized_href(href)
{
	href = href.toLowerCase();
	var slash = href.lastIndexOf("/");
	if (-1 != slash)
	{
		var filename = href.substr(slash + 1);

		var qmark = filename.indexOf("?"); // *** remove string after ?
		if (-1 != qmark) filename = filename.substr(0, qmark); // ***

		var sharp = filename.indexOf("#"); //*** remove string after #
		if (-1 != sharp) filename = filename.substr(0, sharp); // ***

		if ("default.asp" == filename || "index.php" == filename || "default.aspx" == filename)
			href = href.substr(0, slash + 1);
		else
			href = filename;
	}

	return href;
}

function event_onload()
{
	if (!(global.getElementsByTagName || global.all))
		return;

	var items
	if (global.all) 	
		items = global.all.tags("TD")
	else
		items = global.getElementsByTagName("TD");	
		
	var i	
	var lhref = normalized_href(location.href)

	for (i=0; i<items.length; i++)
	{
		var item = items[i]
		if (item.className == "flyoutLink")
		{
			var disabled = false
			var anchors
			if (item.all)
				anchors = item.all.tags("A")
			else
				anchors = item.getElementsByTagName("A")				
			if (anchors.length > 0)
			{
				var anchor = anchors.item(0)
				var ahref = normalized_href(anchor.href)				
				if (ahref == lhref)
				{
					anchor.outerHTML = anchor.innerHTML
					item.style.borderColor = MENU_BORDER_COLOR
					item.style.backgroundColor = MENU_CURRENTPAGE_COLOR
					item.style.cursor = 'default'
					disabled = true					
				}				
			}
			item.defaultBorder = item.style.borderColor
			item.defaultBackground = item.style.backgroundColor
			if (window.addEventListener) { // Mozilla, Netscape, Firefox
				item.addEventListener("mouseover", item_onmouseover, false)
				item.addEventListener("mouseout", item_onmouseout, false)
				if (!disabled) {
					item.addEventListener("mousedown", item_onmousedown, false)
					item.addEventListener("mouseup", item_onmouseup, false)
				}	
			} else { // IE
				item.attachEvent("onmouseover", item_onmouseover)
				item.attachEvent("onmouseout", item_onmouseout)
				if (!disabled) {
					item.attachEvent("onmousedown", item_onmousedown)
					item.attachEvent("onmouseup", item_onmouseup)
				}
			}				
		}
	}	
}



function item_onmouseover(evt)
{	
	var e = whichItem(evt)
	if (document.all) {
		if (e.contains(window.event.fromElement))
			return;
	} else if (document.getElementById) {		
		e.contains = function(node) { 
			var range = document.createRange();
			range.selectNode(this);
			return range.compareNode(node);
		}
		if (e.contains(evt.relatedTarget)==3) // NODE_INSIDE 
		 return;
	} 

	if (e.style.backgroundColor != MENU_CURRENTPAGE_COLOR)
	{		
		e.style.borderColor = MENU_BORDER_COLOR
		e.style.backgroundColor = MENU_MOUSEOVER_COLOR
	}	
	var a
	if (document.all)
		a = e.all.tags("A")
	else if (e.getElementsByTagName)
		a = e.getElementsByTagName("A");
			 
	if (a.length > 0)
		window.status = a[0].href
}

function item_onmouseout(evt)
{	
	var e = whichItem(evt)
	if (document.all) {
		if (window.event.toElement && e.contains(window.event.toElement))
			return; 
	} else if (document.getElementById) {		
		e.contains = function(node) { 
			var range = document.createRange();
			range.selectNode(this);
			return range.compareNode(node);
		}
		if (e.contains(evt.relatedTarget)==3) // NODE_INSIDE 
		 return;
	}	
	e.style.borderColor = e.defaultBorder
	e.style.backgroundColor = e.defaultBackground	
	window.status = ""
}

function whichItem(evt)
{
	var e
	if (document.all) {
		e = event.srcElement;
		while (e.tagName != "TD")
			e = e.parentElement;
	}	else if (document.getElementById) {
		e = evt.target;		
		while (e.tagName != "TD")
			e = e.parentNode;
	}
	return e
}

function item_onmousedown(evt)
{	
	if (document.all) { 
		if (event.button != 1)
			return;
	} else if (document.getElementById) {
		if (evt.which != 1)
			return;
	}
	var e = whichItem(evt)
	e.style.backgroundColor = MENU_MOUSEDOWN_COLOR
	e.mouseIsDown = 1
}

function item_onmouseup(evt)
{	
	if (document.all) { 
		if (event.button != 1)
			return;
	} else if (document.getElementById) {
		if (evt.which != 1)
			return;
	}	
	var e = whichItem(evt)
	if (e.mouseIsDown != 1)
		return	
	e.mouseIsDown = false
	e.style.backgroundColor = MENU_MOUSEOVER_COLOR
	var a
	if (document.all)
		a = e.all.tags("A")
	else if (document.getElementById)		
		a = e.getElementsByTagName("A");	
	if (a.length > 0)
		location.href = a[0].href
}
