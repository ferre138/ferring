<?php
session_save_path('anodnbe');
if (session_id() == "") session_start(); // Initialize Session data
ob_start();
?>
<!doctype html>
<html><!-- InstanceBegin template="/Templates/template.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<meta charset="UTF-8">
<link rel="shortcut icon" href="_images/favicon.ico" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
    <!-- InstanceBeginEditable name="Metas" -->
    <meta name="Description" content="Want to know more about Tuzen® and how it has been clinically proven to eliminate or reduce 95% of the symptoms associated with Irritable Bowel Syndrome?"/>
    <!-- InstanceEndEditable -->
	<!-- InstanceBeginEditable name="doctitle" -->
      <title>Tuzen® | Contact Us</title>
    <!-- InstanceEndEditable -->


<!-- Bootstrap Framework-->
    <link href="_css/bootstrap.css" rel="stylesheet">
    <!-- Styles specific to this project -->
    <link href="_css/theme.css" rel="stylesheet">
    <!-- HTML5 Shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
      <script src="https://oss.maxcdn.com/libs/respond.js/1.4.2/respond.min.js"></script>
    <![endif]-->
    
<!-- InstanceBeginEditable name="head" -->
<!-- InstanceEndEditable -->

	<link href='https://fonts.googleapis.com/css?family=Open+Sans:400,300italic,300,400italic,600,600italic,700,700italic,800,800italic' rel='stylesheet' type='text/css'>
    
</head>

<body>
<div id="wrapper">

<!-- Header -->
<!-- InstanceBeginEditable name="TopLinks" -->
<div id="topnav-wrapper">
  <div class="container">
    <div class="col-sm-12 col-md-12">
      <ul class="nav nav-pills">
        <li><a href="https://www.tuzen.ca/">Home</a></li>
        <li class="lineSpace">|</li>
        <li><a href="French/contact-us.php">Français</a></li>
      </ul>
    </div>
  </div>
</div>
<!-- InstanceEndEditable -->
<!--end topnav-wrapper-->


<div class="container" id="main-nav">
 <div class="row">
  <div class="col-sm-12 col-md-12">
  <a href="https://www.tuzen.ca/"><img src="_images/TUZEN-Logos_RGB_E.jpg" width="195" height="90" id="logo" alt="Tuzen"/></a>
  
<nav class="navbar navbar-default navbar-right" role="navigation">
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
      <ul class="nav navbar-nav">
        <li class="dropdown">
          <a href="about-tuzen.html" class="dropdown-toggle disabled" data-toggle="dropdown">About Tuzen<sup>&reg;</sup></a>
          <ul class="dropdown-menu">
            <li><a href="product-information.html">Product Information</a></li>
            <li class="divider"></li>
            <li><a href="frequently-asked-questions.html">FAQs</a></li>
          </ul>
        </li>
        <li><a href="probiotics-explained.html">Probiotics Explained</a></li>
        <li class="dropdown"><a href="living-with-IBS.html" class="dropdown-toggle disabled" data-toggle="dropdown">Living with IBS</a>
        <ul class="dropdown-menu">
        	<li><a href="IBS-facts.html">IBS Facts</a></li>
            <li class="divider"></li>
           <li><a href="symptoms-of-IBS.html">Symptoms</a></li>
           <li class="divider"></li>
           <li><a href="ibs-management.html">Management</a></li>
        </ul>
        </li>
        <li><a href="clinical-studies.html">Clinical Studies</a></li>
        <li><a href="https://tuzen.ca/store-locator.html" class="buynow">Where To Buy</a></li>
      </ul>
    </div><!-- /.navbar-collapse -->
  </div><!-- /.container-fluid -->
</nav>

  </div>
 </div>
</div>

<!-- header end -->

<!-- InstanceBeginEditable name="mainContent" -->

<div class="top-cont">
  <div class="container">
	<div class="row text-center">
    
      <div class="col-sm-12 col-md-12">
      
      	<h1 class="h1-subtitle">Contact Tuzen<sup>&reg;</sup></h1>
        <h3>Yes, I would like to know more about Tuzen and how it has been clinically proven to eliminate or reduce 95% of the symptoms associated with Irritable Bowel Syndrome.</h3>

      </div>
      
	</div>
  </div>
</div>

<div class="spacer">&nbsp;</div>


<div class="default-wrapper">
  <div class="container">
 	<div class="row">
  <div class="col-md-2">&nbsp;</div>

 <div class="col-sm-12 col-md-8 content">
      
      
      	<?php
			if(isset($_GET['error']) && $_GET['error'] == 'errEmpty'){
				echo '<h3 style="color:red;">All required fields must be filled out.<br><br></h3>';	
			}
		?>
        <?php if(isset($_REQUEST['err']) && $_REQUEST['err']=="wc"){ ?><h3 style="color:red;font-size:14px;">Incorrect verification code.</h3><?php } ?>
        
        <form id="contact-form" method="post" action="inc/mailprocessor.php" onSubmit="return checkFields();">
        <div class="col-sm-6 col-md-6">
         <h6>&nbsp;</h6>
        <h3>
        
		  <div class="form-group">
    	<label for="FirstName">First Name</label>
    	*
    	<input type="text" class="form-control" id="FirstName" name="firstName"  value="<?php if(isset($_SESSION['firstName'])) echo $_SESSION['firstName']; ?>" required>
 		  </div>
        
        <div class="form-group">
   	 	<label for="Company">Organization/Company</label>
    	<input type="text" class="form-control" id="Company" name="company" value="<?php if(isset($_SESSION['company'])) echo $_SESSION['company']; ?>">
  		</div>
        
        <div class="form-group">
   	 	<label for="PostalCode">Postal Code</label>
   	 	*
    	<input type="text" class="form-control" id="PostalCode" name="postal" value="<?php if(isset($_SESSION['postal'])) echo $_SESSION['postal']; ?>" required>
  		</div>
        
        <div class="form-group">
   	 	<label for="Telephone">Telephone*</label>
    	<input type="tel" class="form-control" id="Telephone" name="phone" value="<?php if(isset($_SESSION['phone'])) echo $_SESSION['phone']; ?>" required>
  		</div>
        
        <div class="form-group">
   	 	<label for="Fax">Fax</label>
    	<input type="tel" class="form-control" id="Fax" name="fax" value="<?php if(isset($_SESSION['fax'])) echo $_SESSION['fax']; ?>">
  		</div>
        
        </h3>
       </div><!--End col-md-4-->
       
        <div class="col-sm-6 col-md-6">
        <div style="text-align:right;"><h6>* Mandatory</h6></div>
        <h3>
        
        <div class="form-group">
   	 	<label for="LastName">Last Name</label>
   	 	*
   	 	<input type="text" class="form-control" id="LastName" name="lastName" value="<?php if(isset($_SESSION['lastName'])) echo $_SESSION['lastName']; ?>" required>
  		</div>
        
       <div class="form-group">
   	 	<label for="City">City*</label>
    	<input type="text" class="form-control" id="City" name="city" value="<?php if(isset($_SESSION['city'])) echo $_SESSION['city']; ?>" required>
  		</div>
        
        <div class="form-group">
   	 	<label for="Province">Province</label>
    	<select class="form-control" id="Province" name="province">
 			<option value="">Please Choose..</option>
  			<option value="Ontario" <?php if(isset($_SESSION['province']) && $_SESSION['province'] == 'Ontario') echo 'selected="selected"'; ?>>Ontario</option>
  			<option value="British Columbia" <?php if(isset($_SESSION['province']) && $_SESSION['province'] == 'British Columbia') echo 'selected="selected"'; ?>>British Columbia</option>
  			<option value="Alberta" <?php if(isset($_SESSION['province']) && $_SESSION['province'] == 'Alberta') echo 'selected="selected"'; ?>>Alberta</option>
  			<option value="Saskatchewan" <?php if(isset($_SESSION['province']) && $_SESSION['province'] == 'Saskatchewan') echo 'selected="selected"'; ?>>Saskatchewan</option>
           <option value="Manitoba" <?php if(isset($_SESSION['province']) && $_SESSION['province'] == 'Manitoba') echo 'selected="selected"'; ?>>Manitoba</option>
           <option value="Quebec" <?php if(isset($_SESSION['province']) && $_SESSION['province'] == 'Quebec') echo 'selected="selected"'; ?>>Quebec</option>
           <option value="Newfoundland" <?php if(isset($_SESSION['province']) && $_SESSION['province'] == 'Newfoundland') echo 'selected="selected"'; ?>>Newfoundland</option>
           <option value="PEI" <?php if(isset($_SESSION['province']) && $_SESSION['province'] == 'PEI') echo 'selected="selected"'; ?>>PEI</option>
           <option value="Nova Scotia" <?php if(isset($_SESSION['province']) && $_SESSION['province'] == 'Nova Scotia') echo 'selected="selected"'; ?>>Nova Scotia</option>
           <option value="New Brunswick" <?php if(isset($_SESSION['province']) && $_SESSION['province'] == 'New Brunswick') echo 'selected="selected"'; ?>>New Brunswick</option>
           <option value="Nunavut Territory" <?php if(isset($_SESSION['province']) && $_SESSION['province'] == 'Nunavut Territory') echo 'selected="selected"'; ?>>Nunavut Territory</option>
           <option value="Northwest Territory" <?php if(isset($_SESSION['province']) && $_SESSION['province'] == 'Northwest Territory') echo 'selected="selected"'; ?>>Northwest Territory</option>
           <option value="Yukon Territory" <?php if(isset($_SESSION['province']) && $_SESSION['province'] == 'Yukon Territory') echo 'selected="selected"'; ?>>Yukon Territory</option>
		</select>
  		</div>
        
         <div class="form-group">
   	 	<label for="Email">Email*</label>
    	<input type="email" class="form-control" id="Email" name="email" value="<?php if(isset($_SESSION['email'])) echo $_SESSION['email']; ?>" required>
  		</div>
        
        <div class="form-group">
   	 	<label for="Title">Title</label>
    	<select class="form-control" id="Title" name="title">
 			<option value="">Please Choose..</option>
  			<option value="Primary Care Physician" <?php if(isset($_SESSION['title']) && $_SESSION['title'] == 'Primary Care Physician') echo 'selected="selected"'; ?>>Primary Care Physician</option>
  			<option value="Pharmacist" <?php if(isset($_SESSION['title']) && $_SESSION['title'] == 'Pharmacist') echo 'selected="selected"'; ?>>Pharmacist</option>
  			<option value="Surgeon/Specialist" <?php if(isset($_SESSION['title']) && $_SESSION['title'] == 'Surgeon/Specialist') echo 'selected="selected"'; ?>>Surgeon/Specialist</option>
  			<option value="Caregiver" <?php if(isset($_SESSION['title']) && $_SESSION['title'] == 'Caregiver') echo 'selected="selected"'; ?>>Caregiver</option>
           <option value="Patient" <?php if(isset($_SESSION['title']) && $_SESSION['title'] == 'Patient') echo 'selected="selected"'; ?>>Patient</option>
		</select>
  		</div>
        
        </h3>
        </div><!--End col-md-4-->
        
        
        <div class="col-sm-12 col-md-12">
        <h3>
        <div class="form-group">
        <label for="Message">Reasons for Contact Request*</label>
        <textarea class="form-control" id="Message" name="message" required><?php if(isset($_SESSION['message'])) echo $_SESSION['message']; ?></textarea>
        </div>
        
        <div class="clearfix">&nbsp;</div>
        
        <div class="col-md-12">
        <strong>Please Enter Verification Code:</strong><br><br>
        <a href="#" onclick="document.getElementById('captcha').src = 'securimage/securimage_show.php?' + Math.random(); return false"><img src="securimage/securimage_show.php" width="106" height="50" id="captcha" /></a>&nbsp;<a href="#" onclick="document.getElementById('captcha').src = 'securimage/securimage_show.php?' + Math.random(); return false"><img src="securimage/images/refresh.gif" alt="" style="padding-top:4px;" border="0"/></a><br /><br><input name="captchacode" type="text" class="form-control" id="captchacode" />
        </div>
        
        <div class="clearfix">&nbsp;</div>
        
        <input name="honey" type="text" value="" style="display:none;">
  		<input type="submit" class="purple-btn" value="Submit" />
        </h3>
        </div>
      	</form>
     </div>
	
 <div class="col-md-2">&nbsp;</div>
 
      </div>
  <div class="clearfix">&nbsp;</div>    
    </div> 
 </div>


<div class="spacer">&nbsp;</div>

<!-- InstanceEndEditable -->


<!-- Footer -->
<div id="footer-top">
<div class="container">
 <div class="row "> 
 <div class="col-sm-12 col-md-12">
	<h2 class="white">Feel The Difference</h2>
    	<h4  class="white">For more information, including IBS relief tips on how to get an effective irritable bowel syndrome treatment,
 read our <a href="frequently-asked-questions.html" class="no-color">IBS FAQs</a> or call our toll-free Tuzen® Information Line: <a href="tel:1-800-263-4057" class="tel">1-800-263-4057</a>.</h4>
 </div>
 </div>
</div>
</div><!-- End Footer-Top -->

<br>

<div id="footer-bottom">
<div class="container">
 <div class="row "> 
 <div class="col-sm-12 col-md-12">
	<h6 class="text-center grey"><a href="product-information.html" class="footerLink">Product Information</a> | <a href="http://www.ferring.ca/" class="footerLink" target="_blank">About Ferring</a> | <a href="site-map.html" class="footerLink">Site Map</a> | <a href="legal-disclaimer.html" class="footerLink">Legal Disclaimer</a> | <a href="contact-us.php" class="footerLink">Contact Us</a> | <br class="brHide"><span id="canadian-residents">This website is intended only for Canadian residents.</span></h6>
    
    <br>
    
 	<div class="col-sm-6 col-md-6"> 
  		<img src="_images/ferring-logo.jpg" width="80" id="ferringLogo" alt="Ferring Pharmaceuticals">
			<h6 style="color:#a0a0a0;">Copyright © <script type="text/javascript">var d=new Date();document.write(d.getFullYear());</script>. Ferring Canada. All rights reserved. <br>
    			200 Yorkland Boulevard, Suite 500 North York <br>
    			Ontario Canada M2J 5C1</h6>
	</div>
    
    
      <div class="col-sm-6 col-md-6"><span class="footerDivider">&nbsp;</span>
  	  <h6 class="text-right" style="color:#a0a0a0;">Natural Product Number (NPN) NPN 80019103<br>
Website Design by <a href="http://www.ravenshoegroup.com" target="_blank">Ravenshoe Group</a></h6>
  </div>
    
 </div>
 </div>
</div>
 <div class="col-sm-12 col-md-12">&nbsp;</div>
</div><!--End Footer Bottom-->


</div><!-- wrapper end -->

<!-- Google Analytics -->
<script type="text/javascript">
	
	var _gaq = _gaq || [];
	_gaq.push(['_setAccount', 'UA-1731303-17']);
	_gaq.push(['_trackPageview']);
	
	(function() {
	var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
	ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
	var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
	})();
	
</script>
<!-- end -->

<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
<script src="_js/jquery.min.js"></script>
<!-- Include all compiled plugins (below), or include individual files as needed -->
<script src="_js/bootstrap.min.js"></script>

<script>
$(window).load(function(){
   var width = $(window).width();
   if(width <= 767){
       $('.dropdown-toggle').removeClass('disabled');
   }
   else{
       $('.dropdown-toggle').addClass('disabled');
   }
})
</script>

<!-- InstanceBeginEditable name="footer" -->
<script type="text/javascript" src="_js/validate.js"></script>
<!-- InstanceEndEditable -->

</body>
<!-- InstanceEnd --></html>
