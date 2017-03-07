<?
##All that this script does is to access the database and pull out the properties of the 3 saved objects in the Movie and also the value of the User's Comment. You really don't need to worry about security or anything else because all that this script is doing is just displaying information. 

###This first line of Code is to load your database variables.
include ('Include.inc');

#This line converts name to uppercase so the login is not case sensitive.
$Name = strtoupper ($Name);
$Name = ereg_replace("[^A-Za-z0-9 ]", "", $Name); 

###This connects to your database.
mysql_connect($DBhost,$DBuser,$DBpass);
@mysql_select_db("$DBName");

###This is the SQL statement that will get your settings from the Database.
$result = mysql_query("SELECT * FROM $table WHERE Name = '$Name'");

     ##This sets the variables from the Database
     $Comment = mysql_result($result,0,"Comment");
     $Object1 = mysql_result($result,0,"Object1");
     $Object2 = mysql_result($result,0,"Object2");
     $Object3 = mysql_result($result,0,"Object3");

 if ($Comment == "") {
$Comment = "Hello $Name";
}

#This next bit prints out the value's of the Comment and the positions of the three objects whose properties you want to save. This is the line that Flash will read into the movie. 

print "Comment=$Comment&$Object1$Object2$Object3";

    
?>