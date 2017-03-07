<?
############### .::Comments are indicated by the # symbol - you can erase all of these if needed.
############### .::Author: Jeffrey F. Hill
############### .::Website: www.Flash-dB.com
############### .::If you have any questions either post them in the Flashkit Scripting & Backend Message board - or visit Flash-db.com and email me.

############### Begin GuestBook Script #####################################

##The first 3 lines use a regular expression to match a pattern then replace it with nothing.  The only reason for this is so we only allow necessary characters to be entered into the guestbook. This also takes out slashes which are sometimes added in the post headers to make the string friendly.  You can erase or take these lines out if you want.

	$Name = ereg_replace("[^A-Za-z0-9 ]", "", $Name);
	$Email = ereg_replace("[^A-Za-z0-9 \@\.\-\/\']", "", $Email);
	$news = ereg_replace("[^A-Za-z0-9 \@\.\-\/\']", "", $news);

	$Website = eregi_replace("http://", "", $Website);
	$Website = ereg_replace("[^A-Za-z0-9 \@\.\-\/\'\~\:]", "", $Website);

	$Name = stripslashes($Name);
	$Email = stripslashes($Email);
	$Website = stripslashes($Website);
	$news = stripslashes($news);

####################################################################################
########### Reading and Writing the new data to the GuestBook Database #############

if ($Submit == "Yes") {
#Next line tells the script which Text file to open.
	$filename = "news.txt";

#Opens up the file declared above for reading 

	$fp = fopen( $filename,"r"); 
	$OldData = fread($fp, 80000); 
	fclose( $fp ); 

#Gets the current Date of when the entry was submitted
	$Today = (date ("l dS of F Y ( h:i:s A )",time()));

#Puts the recently added data into html format that can be read into the Flash Movie.

	$Input = "Name: <b>$Name</b><br>Email: <b><u><a href=\"mailto:$Email\">$Email</a></b></u><br>Website: <b><u><a href=\"http://$Website\" target=\"_blank\">$Website</a></b></u><br>news: <b>$news</b><br><i><font size=\"-1\">Date: $Today</font><br><br>.:::.";

#This Line adds the 'NewNews=' part to the front of the data that is stored in the text file.  This is important because without this the Flash movie would not be able to assign the variable 'NewNews' to the value that is located in this text file 

	$New = "$Input$OldData";

#Opens and writes the file.

	$fp = fopen( $filename,"w+"); 
	fwrite($fp, $New, 80000); 
	fclose( $fp ); 
}
####################################################################################
########## Formatting and Printing the Data from the Guestbook to the Flash Movie ##

#Next line tells the script which Text file to open.
	$filename = "news.txt";

#Opens up the file declared above for reading 

	$fp = fopen( $filename,"r"); 
	$Data = fread($fp, 80000); 
	fclose( $fp );

#Splits the Old data into an array anytime it finds the pattern .:::.
	$DataArray = split (".:::.", $Data);

#Counts the Number of entries in the NewNews
	$NumEntries = count($DataArray) - 1;

	print "&TotalEntries=$NumEntries&NumLow=$NumLow&NumHigh=$NumHigh&NewNews=";
	for ($n = $NumLow; $n < $NumHigh; $n++) {
	print $DataArray[$n];
		if (!$DataArray[$n]) {
		Print "<br><br><b>No More entries</b>";
		exit;
		}
	}
	


####################################################################################
###############  End NewNews Script
?>