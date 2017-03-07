
<?

/*        Flash Chit Chat.                                                        */
/*        I special thanks goes out to joost, and Nicola Delbono        */
/*        joost is at <joost@extrapink.com> */
/*        Original script by Nicola Delbono <key5@key5.com>        */
/*        To see it live goto http://www.microcyb.com/chat2/chat */

header("Expires: ".gmdate("D, d M Y H:i:s")."GMT");
header("Cache-Control: no-cache, must-revalidate");
header("Pragma: no-cache");

?>

<?

$person = str_replace ("\n"," ", $person);
$person = str_replace ("<", " ", $person);
$person = str_replace (">", " ", $person);
$person = stripslashes ($person);

?>

&output=

<?
/*        Change this to your filepath. The chat will not work if this                                */
/*        is not changed.                                                                        */

$chat_file_ok = "msg.txt";










/*        $chat_lenght is the number of messages displayed. (Optional, see Readme.txt)                */

$chat_lenght = 13;

/*        $max_file_size is the maximum file size the msg.txt file can reach                */
/*        assuming that any chatter doesn't write a message longer than                */
/*        $max_single_msg_lenght (Optional, see Readme.txt)        */

$max_single_msg_lenght = 100000;
$max_file_size = $chat_lenght * $max_single_msg_lenght;



/* ANYTHING BELOW THIS DOES NOT NEED TO BE MODYFIED        */

$file_size= filesize($chat_file);

/*        if file size is more than allowed then                                                        */
/*                        reads last $chat_lenght messages (last lines of msg.txt file)        */
/*                        and stores them in $lines array                                                */
/*                        then deletes the "old" msg.txt  file and create a new msg.txt        */
/*                        pushing the "old" messages stored in $lines array into the        */
/*                        "new" msg.txt file using $msg_old.                                           */
/*                Note: this is done in order to avoid huge msg.txt file size.                */
                        
if ($file_size > $max_file_size) {

/* reads file and stores each line $lines' array elements        */

$lines = file($chat_file_ok);
/*get number of lines                                                                */

$a = count($lines);

$u = $a - $chat_lenght;
for($i = $a; $i >= $u ;$i--){
                $msg_old =  $lines[$i] . $msg_old;
        }
$deleted = unlink($chat_file_ok);
$fp = fopen($chat_file_ok, "a+");
$fw = fwrite($fp, $msg_old);
fclose($fp);
}

/* the following is because every message has to be                */
/* placed into one single line in the msg.txt file.                        */
/* You can render \n (new lines) with "<br>" html tag anyway.        */

$msg = str_replace ("\n"," ", $message);

/*        if the user writes something...                                        */
/*                the new message is appended to the msg.txt file        */
/*        REMEMBER: the message is appended, hence, if         */
/*                you want the last message to be displayed as the        */
/*                 first one, you have to                                         */

/*                1. store the lines (messages) into the array                */
/*                2. read the array in reverse order                                */
/*                3. post the messages in the output file (the chat)        */
                
/* I added these three lines in order to avoid buggy html code and slashes        */
$msg = str_replace ("\n"," ", $message);
$msg = str_replace ("<", " ", $msg);
$msg = str_replace (">", " ", $msg);
$msg = stripslashes ($msg);




if ($msg != ""){
$fp = fopen($chat_file_ok, "a+");
$fw = fwrite($fp, "$person : $msg\n");
fclose($fp);}

$lines = file($chat_file_ok);
$a = count($lines);

$u = $a - $chat_lenght;

/*        reads the array in reverse order and outputs to chat        */
for($i = $a; $i >= $u ;$i--){
                echo $lines[$i];
        }


?>

