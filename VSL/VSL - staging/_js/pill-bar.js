$(function() {
    function ratingEnable() {

        $('#example-pill').barrating('show', {
            theme: 'bars-pill',
            initialRating: 'A',
            showValues: true,
            showSelectedRating: false,
            onSelect:function(value, text) {
                //alert('Selected rating: ' + value);
                if(value == 9 || value == 10){
                    $('#what_do_you_like_container').show();
                    $('#for_even_higher_rate_container').hide();
                    $('#for_higher_rate_container').hide();
                }else if(value == 7 || value == 8){
                    $('#what_do_you_like_container').hide();
                    $('#for_even_higher_rate_container').show();
                    $('#for_higher_rate_container').hide();
                }else if(value == 0 || value == 1 || value == 2 || value == 3 || value == 4 || value == 5 || value == 6){
                    $('#what_do_you_like_container').hide();
                    $('#for_even_higher_rate_container').hide();
                    $('#for_higher_rate_container').show();
                }

            }
        });
   
    }

    ratingEnable();
});
