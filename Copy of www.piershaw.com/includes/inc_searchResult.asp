<% 

If Request.Form("cat") = "directory" then 
response.redirect "../DUdirectory/searchResult.asp?key=" & Request.Form("key")
end if

If Request.Form("cat") = "news" then 
response.redirect "../DUnews/searchResult.asp?key=" & Request.Form("key")
end if

If Request.Form("cat") = "forums" then 
response.redirect "../DUforum/searchResult.asp?key=" & Request.Form("key")
end if

%>
