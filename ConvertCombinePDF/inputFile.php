<?php
session_start();

$greetingColor = "blue";

echo "<p hidden id='colorChoice' style='color: $greetingColor'>$greetingColor</p>";
$_SESSION['RefreshCount'] = $_SESSION['RefreshCount'] + 1;
?>
