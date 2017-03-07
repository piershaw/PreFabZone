I special thanks goes out to joost, and Nicola Delbono

What do you need to edit:

PHP

The PHP script does not need alterations, unless you want the msg.txt
somewhere else. If you move the msg.txt use an absolute filepath like
/home/account/httpd/chat/msg.txt
>>optional:
The $chat_lenght is the number of lines that are shown in the chat.
Keeping it as low as possible is better for performance. If you are
using different typography or you want to add scrollbars this might
be something to edit.

FLASH

The url to your chat.php3 is set in the pref.txt. Don't forget http://
Change it into something like this: http://www.yourdomain.com/chat/chat.php3
Just edit the url to where your php script is at your server. You can
move and change the chat. Don't change variable names (that includes
textfields!)

INSTALLING

FTP the files to your server. You need to chmod the folder and msg.txt
to 0777. (in telnet > chmod 0777 msg.txt) You can use most popular ftp
programs for this too. The PHP script needs chmod set to: 0755.


RUNNING

I hope it's up and running by now. Just try it in your browser.


