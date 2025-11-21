On Error Resume Next
set arg=wscript.arguments
set fso=createobject("scripting.filesystemobject")
set ws=createobject("wscript.shell")
Set objDialog = CreateObject("UserAccounts.CommonDialog")
do
if arg(0)="" then
objDialog.Filter = "vbs File|*.vbs|All Files|*.*"
objDialog.InitialDir = ""
objDialog.ShowOpen
strLoadFile = objDialog.FileName
Else
strLoadFile=arg(0)
end if
if strLoadFile="" then
k=msgbox("您没有选择任何文件，重新选择吗？",vbYesNo,"vbs代码加密工具")
if k=vbno Then wscript.quit
Else
Exit Do
end if
loop
set f=fso.getfile(strLoadFile)
path=f.parentfolder
name=f.name
set fr=fso.opentextfile(strLoadFile)
dow=13
do while fr.atendofstream=false
line=fr.readline
for i=1 to len(line)
achar=mid(line,i,1)
dow=dow&Chr(44)&asc(achar)
next
dow=dow&chr(44)&"13"&chr(44)&"10"
loop
fr.close
set fw=fso.createtextfile(strLoadFile,2)
fw.write "strs=array("&dow&")"&chr(13)&chr(10)&_
"for i=1 to UBound(strs)"&chr(13)&chr(10)&_
" runner=runner&chr(strs(i))"&chr(13)&chr(10)&_
"next"&chr(13)&chr(10)&_
"Execute runner"