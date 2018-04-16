function checkFields(){
	var error = '';
	
	if($('#name').val() == ''){
		error += 'Name is required.\n';
	}
	if($('#city').val() == ''){
		error += 'City is required.\n';
	}
	if($('#comments').val() == ''){
		error += 'Message is required.\n';
	}
	if (!$("input[name='publish']:checked").val()) {
	  	error += 'Questionnaire is required.\n';
	}
	if(error != ''){
		alert(error);
		return false;	
	}	
	return true;	
}

function checkFieldsLarge(){
	var error = '';
	
	if($('#FirstName').val() == ''){
		error += 'First Name is required.\n';
	}
	if($('#LastName').val() == ''){
		error += 'Last Name is required.\n';
	}
	if($('#Email').val() == ''){
		error += 'Email is required.\n';
	}
	if($('#Telephone').val() == ''){
		error += 'Telephone is required.\n';
	}
	if($('#PostalCode').val() == ''){
		error += 'Postal Code is required.\n';
	}
	if($('#City').val() == ''){
		error += 'City is required.\n';
	}
	if($('#Message').val() == ''){
		error += 'Message is required.\n';
	}
	if(error != ''){
		alert(error);
		return false;	
	}	
	return true;	
}