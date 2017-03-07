<?
// message.php

// open connection to database
$connection = mysql_connect("localhost", "root", "secret") or die
("Unable to connect!");
mysql_select_db("data") or die ("Unable to select database!");

// formulate and execute query
$query = "SELECT message FROM message_table";
$result = mysql_query($query) or die("Error in query: " .
mysql_error());

// get row
$row = mysql_fetch_object($result);

// print output as form-encoded data
echo "msg=" . urlencode($row->FirstName);	  

// close connection
mysql_close($connection);
?>--------------------------------------------------------------------------------