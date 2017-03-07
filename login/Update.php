<?
###This script is used to update the position of each Object in the database.  It combines the x and y position into one column in the database.  

###This first line of Code is to load your database variables.
include ('Include.inc');

###This connects to your database.
mysql_connect($DBhost,$DBuser,$DBpass);
@mysql_select_db("$DBName");

#This is just a simple way to take out unwanted characters.
$Name = ereg_replace("[^A-Za-z0-9 ]", "", $Name); 
$Comment =  ereg_replace("[^A-Za-z0-9 ]", "", $Comment); 

#This line converts name to uppercase so the login is not case sensitive.
$Name = strtoupper ($Name);

###This is the SQL statement that will update the settings in the Database.
$query = "UPDATE saveMovie SET Object1='xO1=$O1x&yO1=$O1y&',  Object2='xO2=$O2x&yO2=$O2y&', Object3='xO3=$O3x&yO3=$O3y' , Comment='$Comment'  WHERE Name='$Name'";
$result = mysql_query($query);


#This prints out the success command to a text field in the movie if it's successful.

print "_root.UpdateStatus=Success Updated - It's saved $Name";
    

?>