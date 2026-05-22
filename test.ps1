$Path = "\\localhost\c$"
$Password = "pass"
$UserName = "user"

# This should just output the arguments
cmd.exe /c echo net.exe use $Path $Password /user:$UserName /persistent:no
