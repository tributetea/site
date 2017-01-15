<?php
/*
Configuration 
You are to edit these configuration values. Not all of them need to be edited.
However, the first few obviously need to be edited.
EMAIL_RECIPIENTS - your email address where you want to get the form submission.

*/

$email_recipients = "signup@tributetea.com";

$visitors_email_field = 'email';//The name of the field where your user enters their email address
                                        //This is handy when you want to reply to your users via email
                                        //The script will set the reply-to header of the email to this email
                                        //Leave blank if there is no email field in your form
$email_subject = "Matcha Mailing-SIGNUP";

$enable_auto_response = true;//Make this false if you donot want auto-response.

//Update the following auto-response to the user
$auto_response_subj = "Sign-Up CONFIRMATION";
$auto_response ="Thank you for subscribing to our Monthly Matcha Mailing.
You have been added to our mailing list.

We look forward to sharing tea news with you soon!

-Your Friends at Tribute Tea
www.tributetea.com";

/*

This is the PHP back-end script that processes the form submission.
It first validates the input and then emails the form submission.
The variable $_POST contains the form submission data.
*/
if(!isset($_POST['submit']))
{
    // note that our submit button's name is 'submit' 
    // We are checking whether submit button is pressed
	// This page should not be accessed directly. Need to submit the form.
	echo "error; you need to submit the form!".print_r($_POST,true);
    exit;
}

require_once "includes/formvalidator.php";
//Setup Validations
$validator = new FormValidator();
$validator->addValidation("email","req","Please fill in Email");
//Now, validate the form
if(false == $validator->ValidateForm())
{
    echo "<B>Validation Errors:</B>";

    $error_hash = $validator->GetErrors();
    foreach($error_hash as $inpname => $inp_err)
    {
        echo "<p>$inpname : $inp_err</p>\n";
    }
    exit;
}

$visitor_email='';
if(!empty($visitors_email_field))
{
    $visitor_email = $_POST[$visitors_email_field];
}

if(empty($email_from))
{
    $host = $_SERVER['SERVER_NAME'];
    $email_from ="Matcha Mailing @ $host";
}

$fieldtable = '';
foreach ($_POST as $field => $value)
{
    if($field == 'submit')
    {
        continue;
    }
    if(is_array($value))
    {
        $value = implode(", ", $value);
    }
    $fieldtable .= "$field: $value\n";
}

$email_body = "You have received a new Matcha Mailing sign-up. Details below:\n$fieldtable\n $extra_info";
    
$headers = "From: $email_from \r\n";
$headers .= "Reply-To: $visitor_email \r\n";
//Send the email!
@mail(/*to*/$email_recipients, $email_subject, $email_body,$headers);

//Now send an auto-response to the user who submitted the form
if($enable_auto_response == true && !empty($visitor_email))
{
    $headers = "From: $email_from \r\n";
    @mail(/*to*/$visitor_email, $auto_response_subj, $auto_response,$headers);
}

//done. 
if (isset($_SERVER['HTTP_X_REQUESTED_WITH'])
    AND strtolower($_SERVER['HTTP_X_REQUESTED_WITH']) === 'xmlhttprequest') 
{
    //This is an ajax form. So we return success as a signal of succesful processing
    echo "success";
}
else
{
    //This is not an ajax form. we redirect the user to a Thank you page
    header('Location: '.$thank_you_url);
}
?>