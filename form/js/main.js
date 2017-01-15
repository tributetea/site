$(document).ready(function () {

    $("#contact-form").validate({
        rules: {
            fullname: {
                required: true
            },
            email: {
                required: true,
                email: true
            },
            message: {
                required: true,
                maxlength: 8000
            }
        },

        messages: { // custom messages
            fullname: {
                required: "Please enter your name"
            },
            email: {
                required: "Please enter your email address"
            },
            message: {
                required: "enter your message",
                maxlength: jQuery.format("The maxlength for message is {0} !")
            },
        },

        submitHandler: function(form) {
            $form = $(form);
            $container = $form.parent();
            w = $form.outerWidth();
            h = $form.outerHeight();
            $form.hide();
            
            $('#msg_submitting', $container).width(w).height(h).fadeIn(1000);            
            $.ajax({
                type: "POST",
                url: $form.attr('action'),
                data: $form.serialize(),
                success: function (data) {
                        $('#msg_submitting', $container).hide();
                        if(data == 'success'){
                            $('#msg_submitted',$container).width(w).height(h).fadeIn(1000);
                        }
                        else{
                            $('#errors',$container).html(data).show();
                            $form.show();
                        }
                }
            });
            
            return false;
        
        }

    });

});