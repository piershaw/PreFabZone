<?php

############################################################
# Help / Support - http://www.flash-db.com/Board/
# Credits Jeff@Flash-db.com
# License - All I'm hoping for is some good looking
# Modifications and usages to this - other then that it's up
# to you. Have fun. - This would all be much easier with MySQL
# But using only text files has got to be good for someone??
############################################################
	// <Start Settings>
	// Split this is what seperates posts on Pages. If you change this you will need to 
	// Start over with new text files.
	$SplitThis = "________________________________________";
	
	// Number of topics to be displayed per page in your Flash movie.
	// It is recommended that you keep this as 11 - it's only experimental if changed.
	$TopicsPerPage = 11;
	// <end settings> 

	$Submit = $HTTP_POST_VARS[Submit];
	$Category = $HTTP_POST_VARS[Category];
	$readDirGlobal = $HTTP_POST_VARS[readDirGlobal];
	$readDir = $HTTP_POST_VARS[readDir];
	$numLow = $HTTP_POST_VARS[NumLow];
	$numHigh = $HTTP_POST_VARS[NumHigh];

	// This just strips special characters from incoming info
	$Category = ereg_replace("[^A-Za-z1-9]", "", $Category);
	$Submit = ereg_replace("[^A-Za-z]", "", $Submit);
	$readDir = ereg_replace("[^A-Za-z]", "", $readDir);
	$readDirGlobal = ereg_replace("[^A-Za-z]", "", $readDirGlobal);

if ($Submit == 'SubmitNew' OR $Submit == 'SubmitReply') {
	$Name = $HTTP_POST_VARS[Name];
	$Subject = $HTTP_POST_VARS[Subject];
	$Message = $HTTP_POST_VARS[Message];
	$File = $HTTP_POST_VARS[File];

	$Name = ereg_replace("[^A-Za-z0-9 ]", "", $Name);
	$Subject = ereg_replace("[^A-Za-z0-9 \.\-\:]", "", $Subject);
	
	$Message = str_replace("-", "%2D", $Message);
	
	$Message = ereg_replace("[^A-Za-z0-9 \@\.\/\%\'\[\]\:\\n\/\r ]", "", $Message);
	$Message = str_replace("%2D", "-", $Message);

	$Name = stripslashes($Name);
	$Subject = stripslashes($Subject);
	$Message = stripslashes($Message);
	$Checked = "Yes";
	
	// Replace URL's and Emails with Links (Using Regular expressions
	// http://www.php.net/manual/en/function.preg-replace.php (for info)
	// All Http address to URL's
	$Message = preg_replace("/([^\w\/])(www\.[a-z0-9\-]+\.[a-z0-9\-]+)/i", "$1http://$2", $Message);
	// All other URL's to URL's
	$Message = preg_replace("/([\w]+:\/\/[\w-?&;#~=\.\/\@]+[\w\/])/i", "<u><A TARGET=\"_blank\" HREF=\"$1\">$1</A></u>", $Message);
	// All Email address's to Links.
	$Message = preg_replace("/([\w-?&;#~=\.\/]+\@(\[?)[a-zA-Z0-9\-\.]+\.([a-zA-Z]{2,3}|[0-9]{1,3})(\]?))/i","<u><A HREF=\"mailto:$1\">$1</A></u>",$Message); 

	// Bad Word Filter For Messages - Just erase the below line if you want people to be able to enter bad words.
	$Message 	= BadWordFunc($Message);
	$Subject 	= BadWordFunc($Subject);
	
	// Tag Replace
	$Message = str_replace("[","<",$Message);
	$Message = str_replace("]",">",$Message);
}

##########################  Script Controls  #####################
// This area just calls different functions dependent on what is passed from the Flash movie.
// Aka - script Controls.

if ($readDir == "Yes") {
	readDirectory($Category, $numLow, $numHigh, $SplitThis, $TopicsPerPage);
}

if ($Submit == 'SubmitNew' AND $Checked == 'Yes') {
	submitNew($Category, $Name, $Subject, $Message, $numLow, $numHigh, $SplitThis, $TopicsPerPage);
}

if ($Submit == 'SubmitReply' AND $Checked == 'Yes') {
	submitReply($Category, $Name, $Subject, $Message, $File, $numLow, $numHigh, $SplitThis, $TopicsPerPage);
}

if ($readDirGlobal == "Yes") {
	$Results1 = CountGlobals("Category1", $SplitThis);
	$Results2 = CountGlobals("Category2", $SplitThis);
	$Results3 = CountGlobals("Category3", $SplitThis);
	$Results4 = CountGlobals("Category4", $SplitThis);

	$GlobalPosts = $Results1[0] + $Results2[0] + $Results3[0] + $Results4[0];
	$GlobalReplys = $Results1[1] + $Results2[1] + $Results3[1] + $Results4[1];

	print "GlobalTopics=$GlobalPosts&GlobalPosts=$GlobalReplys";
}

############  Functions Used ###############################
// The below are the 4 functions used for this message board.


##########  Submit New Post Function #######################
function submitNew($Category, $Name, $Subject, $Message, $numLow, $numHigh, $SplitThis, $TopicsPerPage) {
 
	//Create File Name according to readable format
	$TimePart 		= time();

	$SubjectPart 	= str_replace(" ", "-", $Subject);
	$NamePart 		= str_replace(" ", "-", $Name);
	
	//Put it all Together
	$filename 		= $TimePart."_".$NamePart."_".$SubjectPart.".txt";

	//Get the Current date from Server
	$Today 			= (date ("l dS of F Y ( h:i:s A )",time()));

	//Put the Contents of the Post Together
	$Input 			= "Thread=Subject: <b>$Subject</b><br>Name: <b>$Name</b><br>Comments: $Message<br><i><font size=\"-1\">Date: $Today</font></i><br>$SplitThis";

	//Path to File
	$pathToFile 	= $Category."/".$filename;

	//Write to a New Text File - Category indicates Directory to place file
	$fp = fopen( $pathToFile,"w"); 
	fwrite($fp, $Input, 4000); 
	fclose( $fp ); 

	//Change the Permissions of the file we just created
	chmod ($pathToFile, 0666);

	print "Status=Your file has been added&";
	readDirectory($Category, $numLow, $numHigh, $SplitThis, $TopicsPerPage);
	loadNewReply($pathToFile);
}

###########  Submit Reply Function   ######################

function submitReply($Category, $Name, $Subject, $Message, $File, $numLow, $numHigh, $SplitThis, $TopicsPerPage) {

	$PathName 	= $Category."/".$File;
	$Today 		= (date ("l dS of F Y ( h:i:s A )",time()));

	$newStuff 	= "<br><br>Subject: <b>$Subject</b><br>Name: <b>$Name</b><br>Comments: $Message<br><i><font size=\"-1\">Date: $Today</font></i><br>$SplitThis";

	$fp = fopen( $PathName,"a"); 
	fwrite($fp, $newStuff, 80000); 
	fclose( $fp ); 

	print "Status=Your reply has been posted&";
	readDirectory($Category, $numLow, $numHigh, $SplitThis, $TopicsPerPage);
	loadNewReply($PathName);
}

#########   Read Directory Information Function  ####################

function readDirectory($Category, $numLow, $numHigh, $SplitThis, $TopicsPerPage) {
	$handle = opendir($Category);
		while (false !== ($topics = readdir($handle))) {  
    			if ($topics != "." && $topics != "..") { 
				$topicName = $Category."/".$topics;
    				$topicTime = filemtime($topicName);
    				$topicArray[$topics] = $topicTime;  
			}
		}

	arsort($topicArray);

	$numberOfTopics = sizeOf($topicArray);

	print "numTopicsAll=$numberOfTopics&";

// The rest of this is just getting the topics ordered for each page - Basically it's set at 11/per page as default
// You can change this in the top settings - however it's recommended to leave as is.
// Probably a little easy way of doing this - but anways.
if ($numberOfTopics <= $TopicsPerPage) {
	$Pages = 1;
} 

if ($numberOfTopics > $TopicsPerPage) {
	$Page = ($numberOfTopics / $TopicsPerPage);
	$Pages = floor($Page)+1;
}

$PageNumber = $numHigh / $TopicsPerPage;

if ($numberOfTopics > $numHigh) {
	
	$numDup = (($TopicsPerPage * $PageNumber));

	$numLDup = ($numLow + 1);
} 

if ($numberOfTopics <= $numHigh) {
	$numDup = ($TopicsPerPage-($numHigh-$numberOfTopics));
	$numDup = (($numDup + $numLow));
	$numLDup = ($numLow + 1);
}
	$Count = 1;

	foreach ($topicArray as $key => $value) {

    		$thisTopicName = $key;
    		$thisTopicTime = $value;
    		$thisTopicTime = date("n/j/Y", $thisTopicTime);

		$nameArray 		= split ("_", $thisTopicName);

		$TopicSubject 	= str_replace("-", " ", $nameArray[2]);
		$TopicName 		= str_replace("-", " ", $nameArray[1]);
		$TopicSubject 	= str_replace(".txt", "", $TopicSubject);

    	$topicCreated 	= date("n/j/Y", $nameArray[0]);
		$topicStartedBy = $TopicName;

		$fp = fopen( $Category."/".$thisTopicName,"r"); 
		$numReplysTopics = fread($fp, 80000); 
		fclose( $fp );

		$numReplysTopicsArray = split ($SplitThis, $numReplysTopics);
		$numReplysLocal = count($numReplysTopicsArray) - 1;
		$numReplysGlobal = $numReplysGlobal + $numReplysLocal;

		print "Topic$Count=$TopicSubject&lastModified$Count=$thisTopicTime&File$Count=$thisTopicName&topicCreated$Count=$topicCreated&numReplys$Count=$numReplysLocal&topicStartedBy$Count=$topicStartedBy&";
	$Count = $Count + 1;
		
	}
	closedir ($handle);
	print "TotalPosts=$numReplysGlobal&NumLow=$numLow&NumHigh=$numHigh&Pages=$Pages&numLow=$numLow&numLDup=$numLDup&numDup=$numDup&Go=Yes&";
}

###################  Force Reply to Appear  ###########

function loadNewReply($PathName) {
		$fp = fopen( $PathName,"r"); 
		$ForceIn = fread($fp, 80000); 
		fclose( $fp );

	print "$ForceIn&";
}

###############   Read Global's (Posts and Replys)  #####################

function CountGlobals ($Category, $SplitThis) {

	$handle = opendir($Category);
	while (false !== ($topics = readdir($handle))) {  
    		if ($topics != "." && $topics != "..") { 

    			$topicArray[$topics] = $topicTime;  
		}
	}

	$numberOfTopics = sizeOf($topicArray);

	for($t=0;$t<$numberOfTopics;$t++) {
    		$thisFile = each($topicArray);
    		$thisTopicName = $thisFile[0];

		$fp = fopen( $Category."/".$thisTopicName,"r"); 
		$numReplysTopics = fread($fp, 80000); 
		fclose( $fp );

		$numReplysTopicsArray = split ($SplitThis, $numReplysTopics);
		$numReplysLocal = count($numReplysTopicsArray) - 1;
		$numReplysGlobal = $numReplysGlobal + $numReplysLocal;
	}

	closedir ($handle);
	$DataR = array($numberOfTopics, $numReplysGlobal);
	return $DataR;	
}

// Remove Bad Words Function
function BadWordFunc ($RemoveBadWordText) {
	$RemoveBadWordText = eregi_replace("fuc?k|[kc]unt|asshole|shit|fag|wank|dick|pu[zs]?[zs][yi]|bastard|s[kc]rew|mole[zs]ter|mole[sz]t|coc?k", "****", $RemoveBadWordText);
 	return $RemoveBadWordText;
}
?>