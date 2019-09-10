:: Check where a user is logged in in a domain.
dsquery computer -o rdn > dsquery.computer.txt
for /F %i in (dsquery.computer.txt) do echo %i >> list.txt & qwinsta /SERVER:%i >> list.txt
 
::to kill a session -> rwinsta
