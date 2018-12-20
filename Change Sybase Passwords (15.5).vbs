'!!!NOTE!!! All three server passwords must
'match in order for this script to work

Option Explicit

dim user, cpass, npass
dim s1, s2, s3
dim x

user = inputbox ("Username")
If user = "" Then
    WScript.Quit
End If
cpass = inputbox ("Current password")
If cpass = "" Then
    WScript.Quit
End If
npass = inputbox ("New password")
If npass = "" Then
    WScript.Quit
End If

s1 = "SYBPROD1"
s2 = "SYBPDSS1"
s3 = "SYBDWDSS"

set x = WScript.CreateObject("WScript.Shell")

'CHANGE SYBASE PASSWORDS
x.run("cmd")
WScript.sleep 1000
x.sendkeys "cd C:\sybase\OCS-15_0\bin"
x.sendkeys "~"

x.sendkeys "isql -S " & s1 & " -U " & ucase(user) & " -P " & cpass
x.sendkeys "~"
x.sendkeys "sp_password '" & cpass & "','" & npass & "'"
x.sendkeys "~"
x.sendkeys "go"
x.sendkeys "~"
WScript.sleep 2000
x.sendkeys "exit"
x.sendkeys "~"

x.sendkeys "isql -S " & s2 & " -U " & ucase(user) & " -P " & cpass
x.sendkeys "~"
x.sendkeys "sp_password '" & cpass & "','" & npass & "'"
x.sendkeys "~"
x.sendkeys "go"
x.sendkeys "~"
WScript.sleep 2000
x.sendkeys "exit"
x.sendkeys "~"

x.sendkeys "isql -S " & s3 & " -U " & ucase(user) & " -P " & cpass
x.sendkeys "~"
x.sendkeys "sp_password '" & cpass & "','" & npass & "'"
x.sendkeys "~"
x.sendkeys "go"
x.sendkeys "~"
WScript.sleep 10000
x.sendkeys "exit"
x.sendkeys "~"

Msgbox "Sybase Passwords Changed"

x.sendkeys "exit"
x.sendkeys "~"