Function Convert(MyStr)
			Str = Replace(MyStr, ".", ",")
            Str = -1 * CDbl(Replace(Str, " ", ""))
			Convert = Replace(Str, ",", ".")
End Function

Const RETURNONLYFSDIRS = &H1 
Const NONEWFOLDERBUTTON = &H200 
 
Set oShell = CreateObject("Shell.Application") 
Set oFolder = oShell.BrowseForFolder(&H0&, "Choisir une affaire", RETURNONLYFSDIRS+NONEWFOLDERBUTTON, "x:\") 
If oFolder is Nothing Then  
	MsgBox "Abandon operateur",vbCritical
Else 
  Set oFolderItem = oFolder.Self 
  Racine = oFolderItem.path
reponse = msgbox("Lancer la convertion des fichiers CN?",VbQuestion+VbYesNo, "CN to KKS")
if (reponse = 6) then
Const tssPattern = "nc1"
Const ForReading = 1
Const ForWriting = 2
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FolderExists(Racine+"\convert") Then
'
 Else
  Set objFolder=fso.CreateFolder(Racine+"\convert")
 End If

Set f = fso.GetFolder(Racine+"\")
Set colSubfolders = f.Subfolders
For Each objSubfolder in colSubfolders
set fs = fso.GetFolder(Racine+"\"+objSubfolder.Name)
If fso.FolderExists(Racine+"\convert\"+objSubfolder.Name) Then
'
 Else
 if (objSubfolder.Name<>"convert") then
  Set objFolder=fso.CreateFolder(Racine+"\convert\"+objSubfolder.Name)
  End If
 End If

Set fc = fs.Files
For Each f1 in fc
If Split(f1.name, ".")(1) = tssPattern then 
File = Racine+"\"+objSubfolder.Name+"\"+f1.name
NewFile = Racine+"\convert\"+objSubfolder.Name+"\"+f1.name



rem XS_DSTV_NO_SAWING_ANGLES_FOR_PLATES_NEEDED=FALSE
rem XS_DSTV_CREATE_AK_BLOCK_FOR_ALL_PROFILES = false
 
Set fso = CreateObject("Scripting.FileSystemObject" )
Set objFile = fso.OpenTextFile(File, ForReading)
 
strText = objFile.ReadAll
objFile.Close
tb = split(strText,Chr(10)) 
strTextNew = ""
testAk = 0
if(inStr(tb(8),"TQXD")) then
  tb(9)="M" + Chr(10)
  tb(11)="40.000" + Chr(10)
  tb(12)="10.000" + Chr(10)
  tb(13)="1.500" + Chr(10)
  tb(14)="1.500" + Chr(10)
  a = Convert(tb(18)) + Chr(10)
  tb(18) = Convert(tb(20)) + Chr(10)
  tb(19) = a
  b = Convert(tb(19)) + Chr(10)
  tb(20) = Convert(tb(21)) + Chr(10)
  tb(21) = b
  End if
if(inStr(tb(8),"TUBE_S_30-50-AILE20_L")) then
  tb(18) = Convert(tb(18)) + Chr(10)
  tb(19) = Convert(tb(19)) + Chr(10)
  tb(20) = Convert(tb(20)) + Chr(10)
  tb(21) = Convert(tb(21)) + Chr(10)
  End if
if(inStr(tb(8),"JANSEN_NORM_01531")) then
  tb(9)="SO" + Chr(10)
  tb(11)="50.000" + Chr(10)
  tb(12)="30.000" + Chr(10)
  End if
if(inStr(tb(8),"JANSEN_NORM_01570")) then
  tb(9)="SO" + Chr(10)
  tb(11)="50.000" + Chr(10)
  tb(12)="70.000" + Chr(10)
  End if
if(inStr(tb(8),"PARECLOSE")) then
  tb(9)="SO" + Chr(10)
  tb(11)="20.000" + Chr(10)
  tb(12)="13.000" + Chr(10)
  tb(13)="0.000" + Chr(10)
  tb(14)="0.000" + Chr(10)
  End if
 if(inStr(Split(tb(18), ".")(0),"0")) then tb(18)="0" + Chr(10) End if
  if(inStr(Split(tb(19), ".")(0),"0")) then tb(19)="0" + Chr(10) End if
   if(inStr(Split(tb(20), ".")(0),"0")) then tb(20)="0" + Chr(10) End if
    if(inStr(Split(tb(21), ".")(0),"0")) then tb(21)="0" + Chr(10) End if
For i = LBound(tb) to UBound(tb)
  if(inStr(tb(i),"AK")) then
  testAk = 1
  End if
  if(inStr(tb(i),"IK") or inStr(tb(i),"PU") or inStr(tb(i),"KO") or inStr(tb(i),"SC") or inStr(tb(i),"TO") or inStr(tb(i),"UE") or inStr(tb(i),"PR") or inStr(tb(i),"KA") or inStr(tb(i),"EN") or inStr(tb(i),"BO")) then
  testAk = 0
  End if
  if(testAk = 0) then
  strTextNew = strTextNew + tb(i) 
  End if  
next
If fso.FileExists(Newfile) Then
'
else 
Set ObjFile = fso.createtextFile(Newfile)  
objFile.Close
End if
Set objFile = fso.OpenTextFile(NewFile, ForWriting)
objFile.WriteLine strTextNew
objFile.Close


End If
next
next
End If
End If
Set oFolderItem = Nothing 
Set oFolder = Nothing 
Set oShell = Nothing
