On error resume next 'comment it for debug
dim photoday, lnk, sUrlRequest
set FSO=CreateObject ("Scripting.FileSystemObject")
photoday = fso.GetSpecialFolder(2): if right(bingfile,1)<>"\" then photoday=photoday & "\" : photoday = photoday & "35photo.jpg"

sUrlRequest = "http://ru.35photo.ru/rss/photo_day.xml"
Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
oXMLHTTP.Open "GET", sUrlRequest, False
oXMLHTTP.Send
xmlfile=oXMLHTTP.Responsetext
Set oXMLHTTP = Nothing


beg=instr(lcase(xmlfile),"img src=")
ef=instr(lcase(xmlfile),".jpg")
lnk=mid(xmlfile,beg+14,ef-beg-10)

lnk=replace(lnk,"_temp","_main")
lnk=replace(lnk,"/sizes","")
lnk=replace(lnk,"_800n.jpg",".jpg")

Set oXMLHTTP2 = CreateObject("MSXML2.XMLHTTP")
oXMLHTTP2.Open "GET", lnk, False
oXMLHTTP2.Send
Set oADOStream = CreateObject("ADODB.Stream")
oADOStream.Mode = 3 'RW
oADOStream.Type = 1 'Binary
oADOStream.Open
oADOStream.Write oXMLHTTP2.ResponseBody

oADOStream.SaveToFile photoday, 2

Set objWshShell = WScript.CreateObject("Wscript.Shell")
'use OS to set wallpaper
'objWshShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Desktop\Wallpaper", photoday, "REG_SZ"
'objWshShell.Run "%windir%\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", 1, False

'use irfanview if you want
objWshShell.Run Chr(34) & "C:\Program Files\IrfanView\i_view64.exe" & Chr(34) & " " & Chr(34) & photoday & Chr(34) & " /wall=0 /killmesoftly", 1, False 

Set oXMLHTTP2 = Nothing
Set oADOStream = Nothing
Set FSO = Nothing
Set objWshShell = Nothing