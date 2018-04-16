<script type="text/javascript" src="http://maps.google.com/maps/api/js?sensor=true"></script>
<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.7/jquery.min.js"></script>
<script type="text/javascript" src="jquery.ui.map.full.min.js"></script>
<div id="map_canvas" style="width:1200px;height:800px"></div>
<!--#include file="ewcfg9.asp"-->
<!--#include file="Loginsinfo.asp"-->
<!--#include file="aspfn9.asp"-->
<!--#include file="userfn9.asp"-->
<%
Dim strSQL, rst, orderid, totalAmount, counter
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open EW_DB_CONNECTION_STRING

Set rs = Server.CreateObject("ADODB.Recordset")
'Response.Write StatsGraph(200)


	Dim strXML
	sSql = "SELECT TOP 10 Customers.Inv_FirstName, Customers.Inv_LastName,Inv_Address,inv_city, Customers.inv_Province,inv_PostalCode,CustomerID "
	sSql=sSql & " FROM Customers ORDER BY Customers.CustomerID desc; "
	
	rs.open sSql,conn					
	strXML = ""
	
		i=0
	Do While not rs.EOF
		strXML = strXML & "address["  & i &  "] = """ &  rs("Inv_Address") & "," & rs("inv_city") & ","  &  rs("inv_Province") &  "-"  &  rs("inv_PostalCode")   & """; "
		rs.MoveNext	
		i=i+1		
	Loop
	'strXML = strXML & "</graph>"
	rs.close
%>
<script type="text/javascript">

var geocoder = new google.maps.Geocoder();
var address=new Array();
<%=strXML%>
var tAdds;
tAdds=  '	{"markers":[ ';
var al= address.length;
var k=0;
for (i=0; i < (al-1); i++)
{
geocoder.geocode( { 'address': address[i]}, function(results, status) {
if (status == google.maps.GeocoderStatus.OK) {
	var latitude = results[0].geometry.location.lat();
	var longitude = results[0].geometry.location.lng();
	
	tAdds= tAdds + '{ "latitude":'+ latitude +', "longitude":' + longitude +', "title":"Customers", "content":"'+ results[0].formatted_address +'" }';k=k+1;
	if(i<(al)) tAdds= tAdds + ',';
	}
});
} 
k=k+1;
geocoder.geocode( { 'address': address[i]}, function(results, status) {
if (status == google.maps.GeocoderStatus.OK) {
	var latitude = results[0].geometry.location.lat();
	var longitude = results[0].geometry.location.lng();
	
	tAdds= tAdds + '{ "latitude":'+ latitude +', "longitude":' + longitude +', "title":"Angered", "content":"'+ results[0].formatted_address +'" }';
	
	}
	tAdds= tAdds + '	]}';
	alert(tAdds);
				var data;

			$('#map_canvas').gmap().bind('init', function() { 
					data=jQuery.parseJSON(tAdds);
					$.each( data.markers, function(i, marker) {
						$('#map_canvas').gmap('addMarker', { 
							'position': new google.maps.LatLng(marker.latitude, marker.longitude), 
							'bounds': true 
						}).click(function() {
							$('#map_canvas').gmap('openInfoWindow', { 'content': marker.content }, this);
						});
					});
			});
	
});



</script>
