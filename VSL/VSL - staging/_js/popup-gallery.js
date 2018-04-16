// JavaScript Document


/*Initialize popup and gallery*/

$(document).ready(function() {

	$('.image-popup').magnificPopup({
		type: 'image',
		closeOnContentClick: true,
		image: {
			verticalFit: true
		}
	});
	
	$(document).ready(function() {
	$('.popup-youtube, .popup-vimeo, .popup-gmaps').magnificPopup({
		
		type: 'iframe',
		mainClass: 'mfp-fade',
		removalDelay: 160,
		preloader: false,

		fixedContentPos: false
	});
});
	
});
