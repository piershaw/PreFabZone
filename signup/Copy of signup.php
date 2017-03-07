<?

	$name = stripslashes($name);
	$email = stripslashes($email);
	$website = stripslashes($website);
	$password = stripslashes($password);
        $lastname = stripslashes($lastname);

####################################################################################
########### Reading and Writing the new data to the GuestBook Database #############

if ($Submit == "Yes") {
#Next line tells the script which Text file to open.
	$filename = "userlist.txt";

#Opens up the file declared above for reading 

	$fp = fopen( $filename,"r"); 
	$OldData = fread($fp, 80000); 
	fclose( $fp ); 

#Gets the current Date of when the entry was submitted
	$Today = (date ("l dS of F Y ( h:i:s A )",time()));

#Puts the recently added data into html format that can be read into the Flash Movie.

	$Input = "name: <b>$name</b><br>email: <b><u><a href=\"mailto:$email\">$email</a></b></u><br>website: <b><u><a href=\"http://$website\" target=\"_blank\">$website</a></b></u><br>password: <b>$password</b><br><i><font size=\"-1\">Date: $Today</font><br><br>.:::.";

#This Line adds the 'signup=' part to the front of the data that is stored in the text file.  This is important because without this the Flash movie would not be able to assign the variable 'signup' to the value that is located in this text file 

	$New = "$Input$OldData";

#Opens and writes the file.

	$fp = fopen( $filename,"w+"); 
	fwrite($fp, $New, 80000); 
	fclose( $fp ); 
}
####################################################################################
########## Formatting and Printing the Data from the Guestbook to the Flash Movie ##

#Next line tells the script which Text file to open.
	$filename = "userlist.txt";

#Opens up the file declared above for reading 

	$fp = fopen( $filename,"r"); 
	$Data = fread($fp, 80000); 
	fclose( $fp );

#Splits the Old data into an array anytime it finds the pattern .:::.
	$DataArray = split (".:::.", $Data);

#Counts the Number of entries in the signup
	$NumEntries = count($DataArray) - 1;

	print "&TotalEntries=$NumEntries&NumLow=$NumLow&NumHigh=$NumHigh&GuestBook=";
	for ($n = $NumLow; $n < $NumHigh; $n++) {
	print $DataArray[$n];
		if (!$DataArray[$n]) {
		Print "<br><br><b>No More entries</b>";
		exit;
		}
	}
	


####################################################################################
###############  End signup Script
?>