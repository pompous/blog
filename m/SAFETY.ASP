<%
if session("gokingbook")<>"siva" then
response.redirect "login.asp"
response.end
end if
%>