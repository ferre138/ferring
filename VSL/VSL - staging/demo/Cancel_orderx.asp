<!doctype html>
<!--#include file="ewcfg60.asp"-->
<!--#include file="Customersinfo.asp"-->
<!--#include file="aspfn60.asp"-->
<!--#include file="userfn60.asp"-->
<%

orderid = Request.Querystring("token")
custid = Request.Querystring("txt")
if((orderid & "x"="x") or (custid & "x"="x"))then response.redirect "VSLOrderForm.asp"
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open EW_DB_CONNECTION_STRING
orderid= replace(orderid," ","")		
custid= replace(custid," ","")		

strSQL = "UPDATE Orders SET payment_status='Cancelled', payment_date=Now() " 
strSQL = strSQL & " WHERE payment_status='WIP' and orderid =" & orderid & " and Orders.CustomerId=" & custid & " ;"  
 
conn.Execute(strSQL)
'if orderid = Session("orderid") Then conn.Execute(strSQL)

conn.Close ' Close Connection

%>
<html><!-- InstanceBegin template="/Templates/template.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<meta charset="UTF-8">
<link rel="shortcut icon" href="_images/favicon.ico" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
    <!-- InstanceBeginEditable name="Description" -->
    <meta name="Description" content="I finally was diagnosed with IBS. I was using another probiotic, supplements, and diet but nothing was helping me. I finally read some great reviews on VSL#3 and decided to give it a try. After 1 week of using VSL#3, my digestive problems started disappearing. I started being able to digest food and became regular again. After a month of using this product my health is as good as before. This product works. I recommend it to anyone suffering from digestive issues.  A.G. - October 4th, 2013. “Been taking your product for over three years now. 1 pack a day. After having several bowel resections for Crohn's, being hospitalized for 6 months my doctor recommended trying this before going down other drug treatment paths again. So far I have been in remission with no problems and no other drugs.” P.G. -Sept 23, 2013."/>
    <!-- InstanceEndEditable -->
	<!-- InstanceBeginEditable name="doctitle" -->
      <title>VSL#3&reg; - Testimonials | Probiotic Treatment for IBD IBS Sufferers</title>
    <!-- InstanceEndEditable -->


<!-- Bootstrap Framework-->
    <link href="_css/bootstrap.min.css" rel="stylesheet">
    <!-- Styles specific to this project -->
    <link href="_css/theme.min.css" rel="stylesheet">
    <!-- HTML5 Shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="_js/html5shiv.js"></script>
      <script src="_js/respond.min.js"></script>
    <![endif]-->
    
    <!-- Google Analytics -->
	<script type="text/javascript">
        
        var _gaq = _gaq || [];
        _gaq.push(['_setAccount', 'UA-1731303-15']);
        _gaq.push(['_trackPageview']);
        
        (function() {
        var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
        ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
        var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
        })();
        
    </script>
    <!-- end -->
    
<!-- InstanceBeginEditable name="head" -->
<!-- InstanceEndEditable -->

<link href='http://fonts.googleapis.com/css?family=Open+Sans:400,300italic,300,400italic,600,600italic,700,700italic,800,800italic' rel='stylesheet' type='text/css'>
<link href='http://fonts.googleapis.com/css?family=Oswald:400,700,300' rel='stylesheet' type='text/css'>
    
</head>

<body>
<div id="wrapper">

<!-- Header -->
<div id="topnav-wrapper">
<div class="container">
    <div id="logo"><a href="http://vsl3.ca/"><img src="_images/vsl3-logo.png" width="212" alt="VSL#3"></a></div>
    <div id="logoTab"><a href="http://vsl3.ca/"><img src="_images/vsl3-logo.png" width="165" alt="VSL#3"></a></div>
<div class="row ">
 
  <div class="col-sm-12 col-md-12">
	<ul class="nav nav-pills">
     <!-- <li><a href="#" class="topNavigation">Français</a></li>
      <li style="color:#FFF; padding:10px 5px 0px 5px;">•</li>-->
      <li><a href="contact-us.asp" class="topNavigation">Contact Us</a></li>
      <li style="color:#FFF; padding:10px 5px 0px 5px;" class="desktop">•</li>
      <li class="desktop"><a href="http://vsl3.ca/" class="topNavigation">Home</a></li>
	</ul>
  </div>
   
 </div>
</div>
</div><!-- topnav end -->

<div id="navigation-wrapper">   
<div class="container">
<div class="row ">   
  <div class="col-sm-12 col-md-12">
<nav class="navbar navbar-default" role="navigation">
  <div class="container-fluid">
    <!-- Brand and toggle get grouped for better mobile display -->
    <div class="navbar-header">
      <button type="button" class="navbar-toggle" data-toggle="collapse" data-target="#bs-example-navbar-collapse-1">
        <span class="sr-only">Toggle navigation</span>
        <span class="icon-bar"></span>
        <span class="icon-bar"></span>
        <span class="icon-bar"></span>
      </button>
    </div>
    
<!-- Collect the nav links, forms, and other content for toggling -->
    <div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">  
      <ul id="navigation" class="nav navbar-nav navbar-right">
        <!--<li>
          <a href="#" data-toggle="dropdown" class="mainNavigation">About VSL#3<sup>®</sup> <span class="caret"></span></a>
          <ul class="dropdown-menu">
            <li><a href="#">Why</a></li>
            <li><a href="#">Ingredients</a></li>
            <li><a href="#">Instructions</a></li>
          </ul>
        </li>-->
        <li><a href="about-VSL.html" class="mainNavigation">About VSL#3<sup>&reg;</sup></a></li>
        <li><a href="about-IBD.html" class="mainNavigation">About IBD</a></li>
        <li><a href="diet-and-lifestyle.html" class="mainNavigation">Diet & Lifestyle</a></li>
        <li><a href="testimonials.asp" class="mainNavigation">Testimonials</a></li>
        <li><a href="frequently-asked-questions.html" class="mainNavigation">FAQs</a></li>
        <li><a href="research.html" class="mainNavigation">Research</a></li>
        <li ><a href="VSLOrderForm.asp" class="btn buyNavigation">BUY ONLINE</a></li>
      </ul>
    </div><!-- /.navbar-collapse -->
  </div><!-- /.container-fluid -->
</nav>
   
   </div>

 </div>

</div>
</div><!-- navigation end -->

<!-- header end -->

<!-- InstanceBeginEditable name="mainContent" -->

<div class="default-wrapper" style="padding:20px 0px 0px 0px;">
  <div class="container">
	<div class="row">
     
     <div class="col-sm-12 col-md-12 text-center">
      	<h1 class="h1-header">Cancelled..</h1><%
	Dim Security
	Set Security = New cAdvancedSecurity
	if (Not Security.IsLoggedIn()) then%>
                    <a href="login.asp">login</a>
                    <%else%>
                    <a href="Customersedit.asp">Edit account</a> : <a href="changepwd.asp">Change Password</a> : <a href="logout.asp">logout</a>
                    <%end if
	 %>
      </div>
      
      <div class="col-sm-12 col-md-12 text-center">
      	<h2 class="h2-heading" style="padding:0px 0px 20px 0px;">Cancelled transaction.
       
        </h2>
    
      </div>  
      
      <div class="lineBreak text-center"><img src="_images/page-divide.png" width="629" height="36" class="img-responsive"></div>
     

  </div>
 </div> 
</div>
	



  
   




<!-- InstanceEndEditable -->


<!-- Footer -->
<div class="footer-wrapper">
<div class="container">
<div class="row "> 
 
 <div class="col-sm-12 col-md-12 text-center tb-pad">
 		<h6 style="line-height:20px;"><large>Questions? Call 1.800.263.4057 or <br id="brView"><a href="contact-us.asp" class="blue-link">Fill Out This Form</a></large><br>
        <a href="full-Product-information.html" class="footerlink">Full Product Information</a> | <a href="http://www.ferring.ca" target="_blank" class="footerlink">About Ferring</a> | <a href="legal-notice.html" class="footerlink">Legal Notice</a> | <a href="sitemap.html" class="footerlink">Site Map</a> | <a href="letters-to-insurance.html" class="footerlink">Letters to insurance companies for reimbursement</a> | <a href="contact-us.asp" class="footerlink">Contact Us</a><br>
          <span class="bluehighlight">This website is intended only for Canadian residents.</span><br>
		  <span>Natural Product Number NPN 80037590</span></h6>
      </div>
   <div class="col-sm-6 col-md-6"> 
     <div id="ferringSpace">&nbsp;</div>
  <div id="ferringLogo"><img src="_images/ferring-logo.jpg" width="80" alt="Ferring Pharmaceuticals"></div>
  <div id="ferringCopy"><h6>Copyright © 2014. Ferring Canada. All rights reserved. <br>
    200 Yorkland Boulevard, Suite 500 North York <br>
    Ontario Canada M2J 5C1</h6></div>
</div>
  
  <div class="col-sm-6 col-md-6"> 
  	  <div id="ferringSpaceR">&nbsp;</div>
  	  <div id="ferringRight"><h6>VSL#3<sup>®</sup> and The Living Shield are registered trademarks of VSL Pharmaceuticals Inc.<br id="brView">
		VSL#3<sup>®</sup> is a probiotic blend that is intended to be used under the supervision of a doctor.<br id="brView">
		Please consult with your doctor before trying VSL#3<sup>®</sup></h6></div>
  </div>
<div class="clearfix"></div>

 <div class="col-sm-12 col-md-12 text-right rsgpad"><h6>Website Design By <a href="http://www.ravenshoegroup.com" target="_blank" class="blue-link">Ravenshoe Group</a></h6></div>

</div>
</div>
</div><!-- footer end -->



</div><!-- wrapper end -->

<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
<script src="_js/jquery.min.js"></script>
<!-- Include all compiled plugins (below), or include individual files as needed -->
<script src="_js/bootstrap.min.js"></script>

<!-- InstanceBeginEditable name="footer" -->
<script>
$(function() {
  $('a[href*=#]:not([href=#])').click(function() {
    if (location.pathname.replace(/^\//,'') == this.pathname.replace(/^\//,'') && location.hostname == this.hostname) {
      var target = $(this.hash);
      target = target.length ? target : $('[name=' + this.hash.slice(1) +']');
      if (target.length) {
        $('html,body').animate({
          scrollTop: target.offset().top
        }, 1000);
        return false;
      }
    }
  });
});
</script>
<script type="text/javascript" src="_js/validate.js"></script>

<!-- InstanceEndEditable -->

</body>
<!-- InstanceEnd --></html>
