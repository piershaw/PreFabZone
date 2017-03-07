<?
##The registration script adds a record to the SQL database with the user's Name. That's about it.

###This first line of Code is to load your database variables.
include ('Include.inc');

##This line of Code changes the name to all UPPERCASE. This is so that the login Name is not ##case sensitive. You can make it case sensitive by leaving this line out.

$RegName = strtoupper ($RegName);

##This line takes out everything from the $Name variable that is not a captital or lowercase letter or a interger between 0 and 9.  This line is only for security purposes so that users can not enter anything that could disrupt the database. You can take this line out if you want.  Or you can allow users to enter other symbols by including a \ followed by that character right after.
$RegName = ereg_replace("[^A-Za-z0-9 ]", "", $RegName); 

#Connects to the Database.
$Connect = mysql_connect($DBhost, $DBuser, $DBpass);
@mysql_select_db("$DBName");

#Preforms the SQL query
$query = "INSERT INTO $table (Name, Object1, Object2, Object3, Comment) VALUES ('$RegName' , '', '', '', 'Hello - Edit This')";
$result = mysql_query($query);

#Gets the number of rows affected by the query as a check.
$numR = mysql_affected_rows($Connect);

if ($numR == 0) {
print "Status=Failure Please Fill out all Fields - Register Again";
}
else if ($numR == 1) {
print "Status=Success You can Now Login - Login";
}

else { print "Status=General Error - UserName already in Use";
}


?>