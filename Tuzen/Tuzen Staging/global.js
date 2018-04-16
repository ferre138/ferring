/* global */
var ConcertGolf		= {};
var ConcertGolfObj 	= null;

ConcertGolf = jQuery.Class.create({
  init: function(){
	this.ui_init();
  },
  ui_init: function() {
	//$("#option").hide();
	$("#option1").hide();
	$("#option2").hide();
	$("#option3").hide();
	$("#option4").hide();
	$('#opt1').attr('checked', false);
	$('input[name=2opt1]').attr('checked', false);	
	$('input[name=3opt1]').attr('checked', false);
	$('input[name=4opt1]').attr('checked', false);
	// $('#2opt1').attr('checked', false);
	// $('#3opt1').attr('checked', false);
	// $('#4opt1').attr('checked', false);	
  },
  resetView: function(){
    
    $("#option").show();
	//$("#message").hide();
	// $("#opt1").attr('disabled', false);
	// $("#opt1no").attr('disabled', false);	
  },
  ShowYes: function() {
	$("#option2").hide();
	$("#option1").show();	
	$('input[name=2opt1]').attr('checked', false);	
	$('input[name=3opt1]').attr('checked', false);
	$('input[name=4opt1]').attr('checked', false);
	// $('#2opt1').attr('checked', false);
	// $('#3opt1').attr('checked', false);
	// $('#4opt1').attr('checked', false);			
  },
  ShowNo: function() {
	$("#option2").show();	
	$("#option1").hide();
	$("#option3").hide();		
	$("#option4").hide();
	// $('input[name=opt1]').attr('checked', false);
  },
  ShowNextOption: function() {
    
    if($("#opt1").is(':checked')){
	    if( ($('input[name=2opt1]:checked').val() == "yes") && ($('input[name=3opt1]:checked').val() == "yes")){
		    $("#option4").hide();
			$("#option3").show();			
		}else if( ($('input[name=2opt1]:checked').val() == "yes") && ($('input[name=4opt1]:checked').val() == "yes")){
			$("#option4").hide();
		    $("#option3").show();			
		}else if( ($('input[name=3opt1]:checked').val() == "yes") && ($('input[name=4opt1]:checked').val() == "yes")){
		    $("#option4").hide();
		    $("#option3").show();			
		}else if( ($('input[name=2opt1]:checked').val() == "no") && ($('input[name=3opt1]:checked').val() == "no")){
		    $("#option3").hide();
			$("#option4").show();
		}else if( ($('input[name=2opt1]:checked').val() == "no") && ($('input[name=4opt1]:checked').val() == "no")){
		    $("#option3").hide();
			$("#option4").show();
		}else if( ($('input[name=3opt1]:checked').val() == "no") && ($('input[name=4opt1]:checked').val() == "no")){
		    $("#option3").hide();
			$("#option4").show();
		}else{
			$("#option3").hide();
			$("#option4").hide();
		}
	}
  },
  submit_form: function() {	
	  // this.calculate_total(true);
	  // $("#form_error_display").html('');
	  // var val = this.validate();
	  // if(val !== true) {
		  // $("#form_error_display").html(val);
		  // this._goto_error_top();
		  // return false;
	  // }		  
	  // $("#concert_golf_form input").each(function(i, el) {
		  // if(!$(el).is(':visible'))
			  // $(el).attr('disabled','disabled');
	  // });	
	  // $("#numPlayerInd").attr('disabled',false);	 
 	  // $("#numPlayerSponsor").attr('disabled',false);
 	  // $("#total_amount").attr('disabled',false);	  
	  $("#tuzenform").submit();
	  
  },
  trim: function(str, charlist) {

	    var whitespace, l = 0, i = 0;
	    str += '';
	    
	    if (!charlist) {
	        // default list
	        whitespace = " \n\r\t\f\x0b\xa0\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009\u200a\u200b\u2028\u2029\u3000";
	    } else {
	        // preg_quote custom list
	        charlist += '';
	        whitespace = charlist.replace(/([\[\]\(\)\.\?\/\*\{\}\+\$\^\:])/g, '$1');
	    }
	    
	    l = str.length;
	    for (i = 0; i < l; i++) {
	        if (whitespace.indexOf(str.charAt(i)) === -1) {
	            str = str.substring(i);
	            break;
	        }
	    }
	    
	    l = str.length;
	    for (i = l - 1; i >= 0; i--) {
	        if (whitespace.indexOf(str.charAt(i)) === -1) {
	            str = str.substring(0, i + 1);
	            break;
	        }
	    }
	    
	    return whitespace.indexOf(str.charAt(0)) === -1 ? str : '';
	}
});

// Initialize when doc ready
$().ready(function () {
	ConcertGolfObj = new ConcertGolf();
});