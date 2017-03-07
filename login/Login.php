<?
##The login script accesses the database and checks to see whether your Login name exists. If it does it will select your information from the database and Print it out back to Flash. 

##This first line of Code is to load your database variables - The include File should be ##renamed with your database variables.
include ('Include.inc');

##This line of Code changes the name to all UPPERCASE. This is so that the login Name is not ##case sensitive. You can make it case sensitive by leaving this line out.

$Name = strtoupper ($Name);

##This line takes out everything from the $Name variable that is not a captital or lowercase letter or a interger between 0 and 9.  This line is only for security purposes so that users can not enter anything that could disrupt the database. You can take this line out if you want.  Or you can allow users to enter other symbols by including a \ followed by that character right after.

$Name = ereg_replace("[^A-Za-z0-9 ]", "", $Name); 

#Connects to the database.
mysql_connect($DBhost,$DBuser,$DBpass);
@mysql_select_db("$DBName");

#The SQL query
$query = "SELECT * FROM $table WHERE Name = '$Name'";
$result = mysql_query($query);

#This just gets the number of rows in the Query - It's just a check to see if the Name exists - If not it prints out an error statement.
$numR = mysql_num_rows($result);

#If the number of rows is not equal to one then it prints out an error statement - Name not in Database.    
	if ($numR == 1) { 
	print "Status=Success Login Complete&CheckLog=1";
	}
	else {
	print "Status=Failure - Your name was not found - Please register"; 
	}

?>