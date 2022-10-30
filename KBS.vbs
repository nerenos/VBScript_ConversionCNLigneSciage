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
	If (reponse = 6) Then
		Const tssPattern = "nc1"
		Const ForReading = 1
		Const ForWriting = 2
		Set Fso = CreateObject("Scripting.FileSystemObject")
		If Fso.FolderExists(Racine+"\convert") Then
			'
		Else
			Set objFolder = Fso.CreateFolder(Racine+"\convert")
		End If

		Set f = Fso.GetFolder(Racine+"\")
		Set colSubfolders = f.Subfolders
		For Each objSubfolder in colSubfolders
			Set fs = Fso.GetFolder(Racine+"\"+objSubfolder.Name)
			If Fso.FolderExists(Racine+"\convert\"+objSubfolder.Name) Then
				'
			Else
				If (objSubfolder.Name<>"convert") Then Set objFolder = Fso.CreateFolder(Racine+"\convert\"+objSubfolder.Name) End If
			End If

			Set fc = fs.Files
			For Each f1 in fc
				If Split(f1.name, ".")(1) = tssPattern Then 
					File = Racine+"\"+objSubfolder.Name+"\"+f1.name
					NewFile = Racine+"\convert\"+objSubfolder.Name+"\"+f1.name

					rem XS_DSTV_NO_SAWING_ANGLES_FOR_PLATES_NEEDED=FALSE
					rem XS_DSTV_CREATE_AK_BLOCK_FOR_ALL_PROFILES = false

					Set Fso = CreateObject("Scripting.FileSystemObject" )
					Set objFile = Fso.OpenTextFile(File, ForReading)

					strText = objFile.ReadAll
					objFile.Close
					tb = split(strText,Chr(10)) 
					strTextNew = ""
					testAk = 0
					
					If(inStr(tb(8),"TUBE_S_30-50-AILE20_L")) Then
						tb(18) = Convert(tb(18)) + Chr(10)
						tb(19) = Convert(tb(19)) + Chr(10)
						tb(20) = Convert(tb(20)) + Chr(10)
						tb(21) = Convert(tb(21)) + Chr(10)
					End If
					If(inStr(tb(8),"JANSEN_NORM_01531")) Then
						tb(9)="SO" + Chr(10)
						tb(11)="50.000" + Chr(10)
						tb(12)="30.000" + Chr(10)
					End If
					If(inStr(tb(8),"JANSEN_NORM_01570")) Then
						tb(9)="SO" + Chr(10)
						tb(11)="50.000" + Chr(10)
						tb(12)="70.000" + Chr(10)
					End If
					If(inStr(tb(8),"PARECLOSE")) Then
						tb(9)="SO" + Chr(10)
						tb(11)="20.000" + Chr(10)
						tb(12)="13.000" + Chr(10)
						tb(13)="0.000" + Chr(10)
						tb(14)="0.000" + Chr(10)
					End If
					
					If(inStr(Split(tb(18), ".")(0),"0")) Then tb(18)="0" + Chr(10) End If
					If(inStr(Split(tb(19), ".")(0),"0")) Then tb(19)="0" + Chr(10) End If
					If(inStr(Split(tb(20), ".")(0),"0")) Then tb(20)="0" + Chr(10) End If
					If(inStr(Split(tb(21), ".")(0),"0")) Then tb(21)="0" + Chr(10) End If
					
					For i = LBound(tb) to UBound(tb)
						If(inStr(tb(i),"AK")) Then testAk = 1 End If
						If(inStr(tb(i),"IK") or inStr(tb(i),"PU") or inStr(tb(i),"KO") or inStr(tb(i),"SC") or inStr(tb(i),"TO") or inStr(tb(i),"UE") or inStr(tb(i),"PR") or inStr(tb(i),"KA") or inStr(tb(i),"EN") or inStr(tb(i),"BO")) Then testAk = 0 End If
						If(testAk = 0) Then strTextNew = strTextNew + tb(i) End If  
					Next
					
					If Fso.FileExists(Newfile) Then
						'
					else 
						Set ObjFile = Fso.createtextFile(Newfile)  
						objFile.Close
					End If
					
					Set objFile = Fso.OpenTextFile(NewFile, ForWriting)
					objFile.WriteLine strTextNew
					objFile.Close
				End If
			Next
		Next
	End If
End If

Set oFolderItem = Nothing 
Set oFolder = Nothing 
Set oShell = Nothing
