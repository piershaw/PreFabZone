Creating a GuestBook in Flash
-Jeffrey F. Hill
www.flash-db.com

################################

Theory:

This Guest Book works by storing user entered information in a text file.  The script that is used organizes the data according to when it was entered, in a way that the newest entries are always first.  All of the entries are separated by a pattern that the script later splits into an array.  The array is then manipulated according to commands sent from the Flash movie.  In this way we can organize the data that is returned to the flash movie.  The script returns a fixed amount of data back to the flash movie at one time (default is 10 entries), making the time it takes to view the entries minimal.  Their are quite a few modifications that can be added by further manipulating these array's.  Using this technique our simple Text file becomes a database that we can organize, sort, and filter.  Best of all it will work on any server as long as you can run some type of server side script and can change the permissions of a text file.


################################

Part III - Installing the Script

1) Create a directory called 'GuestBook' - or something similar on your server.

2) Upload the PHP script (GuestBook.php), the swf (GuestBook.swf), and the text file (GuestBook.txt) to this directory. Remember the GuestBook.txt file is initially empty.

3) Change the permissions of the 'GuestBook.txt' file to 777.  You can do this by using Telnet or your FTP program.  With WS_ftp you can right click on the file once uploaded and select 'chmod' - Just check all the boxes off.

4) That's it. It should be up and running if you followed those 3 steps.

################################

Part IV - Error Checking

1) Are you sure that you can run php scripts on your server?  Check with your systems administrator to make sure.

2) Have you changed the permissions of the text file 'GuestBook.txt'.  Make sure their set so you can write to them.

3)  Check to see if you have any errors in your scripts.  Open up the script in your browser window by typing in the full url to the script on your server - If you receive an error message, then you've got a small syntax error in your script.  It's probably a missed ';' or a similar error.  Usually this reports a line number along with the error, this should give you a good start on where to find the error.  If their is a blank screen or a success message then the script is working. 

4) Verify all paths used and try to visually track the variables from the beginning to the end of the process


******
-Jeffrey F. Hill
www.Flash-db.com