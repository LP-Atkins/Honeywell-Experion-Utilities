'Script for automatically retrieving descriptions from Honeywell Series C IO points in use from their respective IO blocks and containing control modules.
'Paul Atkins - October 2024.
'v1.0


On Error Resume Next
Script.Timeout = 600

Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim dataFileFullName: dataFileFullName = "C:\Temp\YOUR_REPORT_NAME.TXT"
Dim objFile
Dim IOMarray
Dim i

	Dim p, o
	Dim CardName, IOType, Channels
	Dim Tagname
	Dim aTagname
	Dim tempString
	Dim descString, descString2
	Dim Len1, Len2


'The array of IO modules goes here, for example:
IOMarray = Array("1S01_LB13_AI01","1S01_LB15_AI02","1S01_LB17_AI03","1S01_LB19_AO01","1S01_MB01_DI01","1S01_MB04_DI02","1S01_MB07_DI03","1S01_MB10_DI04","1S01_MB13_DI05","1S01_MB16_DI06","1S01_MB19_DI07","1S01_RB01_DO01","1S01_RB04_DO02","1S01_RB07_DO03","1S01_RB10_DO04","1S01_RB13_DI08","1S01_RB16_DI09","1S01_RB19_DI10")


'Only run once.
If objFSO.FileExists(dataFileFullName) Then Exit Sub


Set objFile = objFSO.OpenTextFile(dataFileFullName, 2 , True) '2=ForWriting

For i = LBound(IOMarray) To UBound(IOMarray)

	IOType = UCASE(Left(Right(IOMarray(i), 4),2))
	If IOType = "DI" Or IOType = "DO" Then
		Channels = 32
	Else
		Channels = 16
		If Instr(IOMarray(i), "AIL") > 0 Then Channels = 64
	End If
	CardName = IOMarray(i)
	
	
	For p = 1 To Channels

		
		Tagname = Server.ParamValue(CardName & ".CHNLNAME." & p)
		aTagname = Split(Tagname, ".")
		tempString = p & ","
		For o = LBound(aTagname) To UBound(aTagName)
			If UBound(aTagname) = 0 Then
				If Instr(aTagname(0), "CHANNEL") > 0 Then			'CHANNEL indicates a free channel, so it should be omitted.
					tempString = tempString & " ,"
				Else
					tempString = tempString & aTagname(o) & ","
				End If
			Else
				tempString = tempString & aTagname(o) & ","
			End If
		Next

		If UBound(aTagname) = 0 Then tempString = tempString & " ,"	'If there is no parameter the column will get skipped.
		
		descString = vbNullString
		descString2 = vbNullString
		
		descString = Server.ParamValue(Tagname & ".DESC")		'Parameter Description
		descString2 = Server.ParamValue(aTagname(0) & ".DESC")	'Point Description.
		descString = Trim(descString)
		descString2 = Trim(descString2)
		descString = Replace(descString, ",", "...")
		descString2 = Replace(descString2, ",", "...")
		
		Len1 = Fix(Len(descString) * 0.80)
		Len2 = Fix(Len(descString2) * 0.80)
		
		If Len1 > 5 And Len2 > 5 Then	'If either of the strings are relatively short, just move on.
			
			'Scrub the secondary description if it is very similiar to the original one.
			If Len(descString) > Len(descString2) Then
				If Left(descString,Len2) = Left(descString2,Len2) Then descString2 = vbNullString
			Else
				If Left(descString,Len1) = Left(descString2,Len1) Then descString2 = vbNullString
			End If
			
			'If it is a shorter secondary or primary string then remove one of them.
			If Instr(1, descString, descString2) Then descString2 = vbNullString
			If Instr(1, descString2, descString) Then descString = vbNullString
			
			If Len(descString) = 0 Then	descString = descString2
			If descString = descString2 Then descString2 = vbNullString
		
		End If
		
		If descString2 = vbNullString Then
			tempString = tempString & descString
			'tempString = tempString & Chr(34) & descString & vbCrLf & descString2 & Chr(34)
		Else
			tempString = tempString & Chr(34) & descString2 & vbCrLf & descString & Chr(34)
		End If
		
		'Get the init request status of outputs.
		If IOType = "AO" Or IOType = "DO" Then
			tempString = tempString & "," & Server.ParamValue(Tagname & ".INITREQ")
		End If
		
		objFile.WriteLine tempString
	
	Next
	
	
	
Next


objFile.Close