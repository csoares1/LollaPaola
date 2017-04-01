<?php
$to = "you@yourdomain.com";
$subject = "Hi!";
$body = "TEST";
$headers = "MIME-Version: 1.0\r\n";
$headers .= "Content-type:text/html;charset=UTF-8\r\n";
$headers .= "From: You <you@yourdomain.com>\r\n";

if(mail($to,$subject,$body,$headers)) {
echo "MAIL - OK";
} else {
echo "MAIL FAILED";
}
?>
