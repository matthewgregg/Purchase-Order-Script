Function strClean (strtoclean)
	Dim objRegExp, outputStr
	Set objRegExp = New Regexp
	objRegExp.IgnoreCase = True
	objRegExp.Global = True
	objRegExp.Pattern = "[/\:*?""<>|]+"
	outputStr = objRegExp.Replace(strtoclean, "")
	strClean = outputStr
End Function
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = WScript.CreateObject ("WScript.Shell")
varPathCurrent = objFSO.GetParentFolderName(WScript.ScriptFullName)
varPathParent = objFSO.GetParentFolderName(varPathCurrent)
strCurrent = varPathCurrent & "\"
strHome = objShell.ExpandEnvironmentStrings(varPathParent & "\")
Set objFolder = objFSO.GetFolder(strCurrent)
Set colFiles = objFolder.Files
PdfCount = 0
For Each objFile in colFiles
	If instr(objFile.Name,".pdf") <> 0 Then
		pdfname = objFile.Name
		PdfCount = PdfCount + 1
	End If
Next
If PdfCount = 0 Then
	Wscript.Quit
ElseIf PdfCount = 1 Then
	objFSO.MoveFile pdfname,"po.pdf"
	objShell.Run "pdftotext.exe -raw po.pdf po.txt",0,True
ElseIf PdfCount > 1 Then
	MsgBox "Found " & PdfCount & " pdf files. Please fix manually",0+48+0+4096,"Multiple pdfs Found"
	Wscript.Quit
End If
Set objPON = objFSO.OpenTextFile("po.txt", 1)
	lineNum = objPON.ReadLine
	strNum = Trim(Mid(lineNum, 17))
	chkNum = Trim(Left(lineNum, 15))
objPON.Close
Set objPOV = objFSO.OpenTextFile("po.txt", 1)
	For i = 1 to 3
    	objPOV.ReadLine
	Next
	lineVendor = objPOV.ReadLine
	strVendor = strClean(Trim(Mid(lineVendor, 5)))
	chkVendor = Trim(Left(lineVendor, 3))
objPOV.Close
If Not chkNum = "Purchase Order:" Or Not chkVendor = "To:" Then
	MsgBox "Bad .pdf format. Please continue manually. The file is the '_Unsorted POs' folder"
	objFSO.MoveFile "po.pdf", strHome & "\Purchase Orders\_Unsorted POs\po.pdf"
	objFSO.DeleteFile "po.txt"
	Wscript.Quit
End If
name = strVendor & " " & strNum & ".pdf"
path = strHome & "Purchase Orders\" & strVendor & "\"
If strVendor = "RS Components" Then
	RS = MsgBox ("The PO '" & strNum & "' is an RS order. Do you want to process it through RS-online? Clicking 'No' will email the PO as normal.",4+32+0+4096,"RS Found")
	If RS = 6 Then
		objShell.Run "pdftotext.exe -table po.pdf po.txt",0,True
		If objFSO.FileExists(path & name) Then
			If objFSO.FileExists("po.pdf") Then
				objFSO.DeleteFile "po.pdf"
			End If
			FileExists = MsgBox ("The PO '" & name & "' already exists. Do you want to continue?",4+32+256+4096,"pdf Exists")
			If FileExists = 7 Then
				If objFSO.FileExists("po.txt") Then
					objFSO.DeleteFile "po.txt"
				End If
				Wscript.Quit
			End If
		Else
			objFSO.MoveFile "po.pdf", path & name
		End If
		objShell.Run "RS.hta", 1, True '"cmd /c %SystemRoot%\System32\mshta.exe " & Chr(34) & varPathCurrent & "\RS.hta" & Chr(34),0,False '64 bit mshta.exe
		WScript.Quit
	End If
End If
If Not objFSO.FolderExists(path) Then
	Do
		NoFolderMsgBox1 = MsgBox ("No folder found for '" & strVendor & "'. Do you want to create a new vendor folder? You can click 'No' to rename an existing folder. Click 'Cancel' to move manually.",3+32+0+4096,"Create New Folder?")
		If NoFolderMsgBox1 = 6 Then
			Set objFolder = objFSO.CreateFolder(path)
		ElseIf NoFolderMsgBox1 = 7 Then
			objShell.Run "explorer.exe /e," & """" & strHome & "Purchase Orders\" & """"
			Do
				NoFolderMsgBox2 = MsgBox ("Rename a folder to '" & strVendor & "' and click 'Retry' to search again.",5+0+4096,"Search for Vendor Folder")
			Loop Until objFSO.FolderExists(path) Or NoFolderMsgBox2 = 2
		ElseIf NoFolderMsgBox1 = 2 Then
			objFSO.MoveFile "po.pdf", strHome & "\Purchase Orders\_Unsorted POs\" & name
			objFSO.DeleteFile "po.txt"
			MsgBox "The PO is in the '_Unsorted POs' folder.",0+64+0+4096,"pdf Moved"
			Wscript.Quit
		End If
	Loop Until objFSO.FolderExists(path)
End If
If objFSO.FileExists(path & name) Then
	If objFSO.FileExists("po.pdf") Then
		objFSO.DeleteFile "po.pdf"
	End If
	FileExists = MsgBox ("The PO '" & name & "' already exists. Do you want to email anyway?",4+32+256+4096,"pdf Exists")
	If FileExists = 7 Then
		If objFSO.FileExists("po.txt") Then
			objFSO.DeleteFile "po.txt"
		End If
		Wscript.Quit
	End If
Else
	objFSO.MoveFile "po.pdf", path & name
End If
If objFSO.FileExists("po.txt") Then
	objFSO.DeleteFile "po.txt"
End If
If Left(strNum,2) = "EN" Then
	MailConfirm = MsgBox ("The PO '" & strNum & "' begins with 'EN'. Do you want to email anyway?",4+32+256+4096,"Create email")
		If MailConfirm = 7 Then
			Wscript.Quit
		End If
End If
Do
UpdateContactTrue = 0
	Set objTextFile = objFSO.OpenTextFile(strHome & "contact.csv", 1)
	ArrayMatch = 0
	Do Until ArrayMatch = 1 Or objTextFile.AtEndOfStream
		strNextLine = objTextFile.Readline
		arrContact = Split(strNextLine,",")
		If arrContact(0) = strVendor Then
			ArrayMatch = 1
		Else
			arrContact(1) = ""
		End If
	Loop
	Do
		If ArrayMatch = 0 Then
			UpdateContact = MsgBox ("No contact information found. Would you like to update the contact list?",4+64+0+4096,"No Contact")
			If UpdateContact = 6 Then
				arrContact = Array(strVendor,"","","")
				Do
					strEnterName = InputBox ("Please enter the name of the contact as 'firstname lastname' e.g. John Smith. This field is not required", "Create New Contact")
					strNameChkTrim = Trim(strEnterName)
					strNameChk = Len(strNameChkTrim) - Len(Replace(strNameChkTrim, " ",""))
				Loop Until strNameChk < 2
				If Not IsEmpty(strEnterName) Then
					strEnterEmail = InputBox ("Please enter the email address of the contact. This field is required", "Create New Contact")
					strEmailChk = Len(Trim(strEnterEmail)) - Len(Replace(Trim(strEnterEmail), "@",""))
					If strEmailChk <> 0 Then
						If strNameChk = 0 Then
							strFirstName = strEnterName
							strLastName = ""
						ElseIf strNameChk = 1 Then
							strFirstName = Split(Trim(strEnterName)," ")(0)
							strLastName = Split(Trim(strEnterName)," ")(Ubound(Split(Trim(strEnterName)," ")))
						End If
						On Error Resume Next
						OpenFile = 0
						Do
							Set ContactWrite = objFSO.OpenTextFile(strHome & "contact.csv",8)
							If Err.Number = 0 Then
								OpenFile = OpenFile + 1
								Err.Clear
							ElseIf Err.Number <> 0 Then
								Err.Clear
							End If
						Loop Until OpenFile = 1
						On Error GoTo 0
						ContactWrite.WriteLine strVendor & "," & Trim(strFirstName) & "," & Trim(strLastName) & "," & Trim(strEnterEmail)
						ContactWrite.Close
						UpdateContactTrue = 1
					End If
				End If
			End If
		End If
	Loop Until UpdateContactTrue = 1 Or UpdateContact = 7 Or ArrayMatch = 1
Loop While UpdateContactTrue = 1
If Trim(arrContact(1)) = "" Then
	SpaceContact = ""
Else
	SpaceContact = " "
End If
Set objOutlook = CreateObject("Outlook.Application")
Set objMail = objOutlook.CreateItem(0)
	signature = objMail.HTMLbody
	strbody = "<BODY style=font-size:11pt;font-family:Calibri;color:#1F497D>Hi" & SpaceContact & Trim(arrContact(1)) & ",<br><br>Please find attached a copy of order " & strNum & ".<br>Can you please confirm receipt of this order?</BODY>"
	strSig = "<BODY style=font-size:10pt;font-family:Calibri;color:#1F497D><br>Kind Regards,</BODY><BODY style=" & Chr(34) & "font-size:16pt;font-family:Blackadder ITC;color:#1F497D" & Chr(34) & ">Insert user name here</BODY><BODY style=font-size:10pt;font-family:Calibri;color:#1F497D>MRO Buyer,<br>L.E. Pritchitt & Co. Ltd. (A Lakeland Dairies Company)<br>Tel: +44(0)2891824817.<br>Mob: +44(0)7798750678.<br>Fax: +44(0)2891824824.</BODY><BODY style=font-size:10pt;font-family:Calibri><a Href=http://www.lakeland.ie>http://www.lakeland.ie</a>"
    With objMail
		If ArrayMatch = 1 Then
			.Recipients.Add Trim(arrContact(3))
		End If
		.Subject = strVendor & " Order " & strNum
		.HTMLBody = strbody & strSig
		.Attachments.Add(path & name)
        '.SendUsingAccount = OutApp.Session.Accounts.Item(1) 'Change Item(1)to the account number that you want to use
		.Display
		'.Send 'Don't uncomment this unless you want the email automatically sent!
		'.Quit
	End With
objShell.AppActivate strVendor & " Order " & strNum
Set objMail = Nothing
Set objOutlook = Nothing
