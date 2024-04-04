<?php
/*time.php
Returns server time and a random display color
in response to an AJAX request
*/
session_start();
if ($_SESSION['timedateRefreshCount'] == 0)
{
    $greetingColor = "black";
}
else
{
  $colorPicker = rand(1, 100);
  if ($colorPicker > 75)
      $greetingColor = "red";
  else if ($colorPicker > 50)
      $greetingColor = "green";
  else if ($colorPicker > 25)
      $greetingColor = "blue";
  else
      $greetingColor = "maroon";
}
$date = date("l, F jS");
$time = date('g:ia');
echo "It's $date.<br>
     The time is $time.";
echo "<p hidden id='colorChoice'>$greetingColor</p>";
$_SESSION['timedateRefreshCount'] = $_SESSION['timedateRefreshCount'] + 1;
?>
