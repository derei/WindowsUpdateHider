' This script has been tested under Windows 7 x64 only.
' It should work fine under other versions of Windows,
' but it should be tested nonetheless.

' In order for this script to run properly, it needs a text file
' located in the same place with this script, named "HideUpdatesList.txt"
' and which should contain the KB ids you want to hide, onle ID per line (eg: KB2505438)


Option Explicit

Call ForceAdmin()

' ##########################################################################################
' #################   Force output in CMD Console instead of message box   #################
' #################  		 and few other helpful functions ... 		   #################
' ##########################################################################################

' force running script with elevated privileges (as Admin)
' and uses cscript.exe as interpreter (to have output in console)
Function ForceAdmin
	Dim ObjShell

	If WScript.Arguments.length = 0 Then
		Set ObjShell = CreateObject("Shell.Application")
		ObjShell.ShellExecute "cscript.exe", """" & WScript.ScriptFullName & """" & " RunAsAdministrator",  "", "runas", 1
		'WScript.Quit
	End If

End Function


' alias for printing out text in Console
Function printf(txt)
	On Error Resume Next
    WScript.StdOut.WriteLine txt
	'Wscript.Echo txt
	On Error Goto 0
End Function


Function readln
	Dim usrInput
	
	On Error Resume Next
	usrInput = WScript.StdIn.ReadLine
	On Error Goto 0
	
	readln = usrInput
End Function


' timer to delay Console closing
Function wait(n)
	printf vbNewLine & "Waiting for " & n & " seconds..."
   WScript.Sleep Int(n * 1000)
End Function

' Halt function until user 
Function PressKeyToContinue(msg)
	printf MSG_PRESS_ENTER & msg
	PressKeyToContinue = readln	
End Function


'returns an ArrayList without duplicates
function removeDuplicates(arraylist)
	Dim element, q
	Dim result : Set result = CreateObject("System.Collections.ArrayList")
	
	For Each element in arrayList
		q = result.Contains(element)
		
		If q = False Then
			result.Add(element)
		End If
		
	 Next
	Set removeDuplicates = result
end function

' ##########################################################################################
' ###################       GLOBAL VARIABLES DECLARATION       #############################
' ##########################################################################################

Dim StartTime, ElapsedTime, wuResult

Dim filename        : filename            = "HideUpdatesList.txt"
Dim WshShell        : Set WshShell        = CreateObject( "WScript.Shell"				 )
Dim notInstUpds		: Set notInstUpds	  = CreateObject( "System.Collections.ArrayList" )
Dim installedUpds   : Set installedUpds	  = CreateObject( "System.Collections.ArrayList" )
Dim updatesToHide	: Set updatesToHide	  = CreateObject( "System.Collections.ArrayList" )
Dim blacklistedKBs  : Set blacklistedKBs  = CreateObject( "System.Collections.ArrayList" )


' ##########################################################################################
' ####################################     MESSAGES     ####################################
' ##########################################################################################

Dim MSG_STARTING_PROGRAM  : MSG_STARTING_PROGRAM  = vbTab & "*****************************************************************"                 & vbNewLine &_
													vbTab & "*" & vbTab & "  This script will search for Windows Updates and " & vbTab & "*" & vbNewLine &_
													vbTab & "*" & vbTab & "hide any matching updates based on the list provided "    & vbTab & "*" & vbNewLine &_
													vbTab & "*" & vbTab & vbTab & "in the file " & filename       & vbTab & vbTab & vbTab & "*" & vbNewLine &_
													vbTab & "*****************************************************************"                 & vbNewLine

Dim MSG_PARSING_TXT_FILE  : MSG_PARSING_TXT_FILE  = "Parsing " & filename & " ... "

Dim MSG_COUNT_FOUND_IDS   : MSG_COUNT_FOUND_IDS   = vbTab & "Found "

Dim MSG_UPDATES_TO_HIDE   : MSG_UPDATES_TO_HIDE   = "FOLLOWING UPDATES ARE READY TO BE HIDDEN: " & vbNewLine

Dim MSG_INSTALLED_UPDATES : MSG_INSTALLED_UPDATES = "The following BLACKLISTED UPDATES are already installed on your System:" & vbNewLine

Dim MSG_UNINSTALL_WARNING : MSG_UNINSTALL_WARNING = "BEWARE THAT UNINSTALLING ANY INSTALLED UPDATE MAY CRASH YOUR SYSTEM !!" &_
													 vbNewLine & "*** Proceed with care! ***" & vbNewLine
													 
Dim MSG_WU_SEARCHING	  : MSG_WU_SEARCHING 	  = "Searching for Windows Updates. This may take a while..." & vbNewLine												

Dim MSG_WU_SEARCH_DONE    : MSG_WU_SEARCH_DONE    = "WINDOWS UPDATE CHECK FINISHED!"

Dim MSG_PRESS_ENTER		  : MSG_PRESS_ENTER		  = "PRESS [ENTER] TO "

DIM MSG_START_WU_SEARCH   : MSG_START_WU_SEARCH   = "Start searching for Windows Updates..."

Dim MSG_SHOW_INST_UPDS    : MSG_SHOW_INST_UPDS    = "Display Blacklisted Updates already installed on your System..."

Dim MSG_SHOW_UPDS_TO_HIDE : MSG_SHOW_UPDS_TO_HIDE = "Display Blacklisted Updates ready to be hidden..."

Dim MSG_HIDE_UPDATES_INFO : MSG_HIDE_UPDATES_INFO = "HIDE THE UPDATES!" & vbNewLine &_
													"Hidden Updates won't show in Windows Update anymore, unless restored."

Dim MSG_HIDE_UPD_SUCCESS      : MSG_HIDE_UPD_SUCCESS      = "Update hidden successfully: "										
													
Dim MSG_EXIT_PROGRAM      : MSG_EXIT_PROGRAM      = " Exit the program."													
													
Dim MSG_CHRONO_START      : MSG_CHRONO_START      = "Start time: "

Dim MSG_CHRONO_END		  : MSG_CHRONO_END        = "Time required: "

Dim MSG_SEPARATOR_LINE    : MSG_SEPARATOR_LINE    = " -------------------------------------------------------------------" & vbNewLine
													
' ##########################################################################################
' #######################       FUNCTIONS DECLARATION       ################################
' ##########################################################################################

' Uses regex matching to check for KB number at the beginning of the provided string
Function ValidateKB(str)
	Dim match, validKB, objRegExp, returnVal
	
	Set objRegExp = New RegExp
	returnVal = Null
	
	With objRegExp
		.Pattern    = "^(KB)[1-9]\d*"
		.IgnoreCase = True
		.Global     = False
	End With
	
	Set match = objRegExp.Execute(str)
	
	If match.Count > 0 Then
		validKB = match(0).Value
		returnVal = validKB
	End If
	
	ValidateKB = returnVal
End Function


' retrieves the KB NUMBER from a Windows Update Object
Function GetKBID(upd)
	Dim k, kbID
	
	For Each k in upd.KBArticleIDs
		kbID = k
	Next
	
	GetKBID = "KB" & kbID
End Function


' Compares a provided KB Number against the KB ID of the provided update
' returns NULL if no match, or KBxxxxxxx if match found
Function MatchUpdate(kbNo, upd)
	Dim kbID, returnVal
	
	kbID = GetKBID(upd)
	
	If kbNo = kbID Then
		returnVal = kbNo
	Else
		returnVal = Null
	End If
	
	MatchUpdate = returnVal
End Function


' Will parse the text file, line by line and will populate <blacklistedKBs>
Function ParseTxtFile(txtfile)
	Dim fso, path, file, line, kb, count
	
	Set fso  = CreateObject("Scripting.FileSystemObject")
	path     = fso.GetParentFolderName(WScript.ScriptFullName)
	Set file = fso.OpenTextFile(path & "/" & txtfile)
	count    = 0
	
	' makes sure the ArrayList is clear
	blacklistedKBs.Clear
	
	printf MSG_PARSING_TXT_FILE
	
	Do Until file.AtEndOfStream
		line  = file.ReadLine
		kb	  = ValidateKB(line)
		
		If NOT(IsNull(kb)) Then
			blacklistedKBs.Add kb
			count = count + 1
		End If
	Loop
	
	file.Close
	printf MSG_COUNT_FOUND_IDS & count & " KB IDs." & vbNewLine
End Function


' Performs online Windows Update check and returns result as SearchResult Object
Function PerformUpdateCheck
	Dim updateSession, updateSearcher, searchResult, ssWindowsUpdate
	
	printf MSG_WU_SEARCHING
	
	ssWindowsUpdate   				= 2
	Set updateSession 				= CreateObject("Microsoft.Update.Session") 
	Set updateSearcher 				= updateSession.CreateupdateSearcher()
	updateSearcher.ServerSelection 	= ssWindowsUpdate
	Set searchResult 				= updateSearcher.Search("Type='Software' and IsHidden=0")
	
	printf MSG_WU_SEARCH_DONE
	
	Set PerformUpdateCheck = searchResult
End Function


' Cycles trough Updates (found by <PerformUpdateCheck>).
' populates the global ArrayLists <installedUpds> and <notInstUpds>
Sub ParseUpdates(updatesResultObj)
	Dim i, foundUpdates, wUpdate
	
	Set foundUpdates = updatesResultObj
	
	' make sure the ArrayLists are reset (empty)
	installedUpds.Clear
	notInstUpds.Clear
	
	For i = 0 to foundUpdates.Updates.Count - 1
	
		Set wUpdate = foundUpdates.Updates.Item(i)
		
		If wUpdate.IsInstalled <> 0 Then
			installedUpds.Add wUpdate
			
		Else
			notInstUpds.Add wUpdate
			
		End If
	Next
End Sub


' formats the Windows Update Information in a intuitive way
' and prints it in console.
Sub PrintUpdateDetails(upd)
	Dim uTitle, uDate, uKBid, msg
	
	uTitle = upd.Title
	uDate  = upd.LastDeploymentChangeTime
	uKBid  = GetKBID(upd)
	msg    = vbTab & uKBid & " - " & uDate & " - " & uTitle
	
	printf msg
End Sub


' Shows blacklisted updates that are already on the machine
' ATTENTION! - UNINSTALLING ALREADY INSTALLED UPDATE CAN CRASH WINDOWS!!
Sub ShowInstalledUpds
	Dim kb, upd, matchedKB
	
	printf MSG_INSTALLED_UPDATES
	
	For Each kb In blacklistedKBs
		For Each upd in installedUpds
		
			matchedKB = MatchUpdate(kb, upd)
			if NOT(IsNull(matchedKB)) Then
				PrintUpdateDetails(upd)
			End If
		Next
	Next
	
	printf vbNewLine
	printf MSG_UNINSTALL_WARNING
	printf MSG_SEPARATOR_LINE
End Sub


Function ShowUpdatesToHide
	Dim kb, upd, matchedKB, uth
	Dim tempUpdToHide : Set tempUpdToHide = CreateObject("System.Collections.ArrayList")
	
	' make sure the ArrayList is reset (empty)
	updatesToHide.Clear
	
	For Each kb in blacklistedKBs
	
		For Each upd in notInstUpds
			matchedKB = MatchUpdate(kb, upd)
			
			if NOT(IsNull(matchedKB)) Then
				tempUpdToHide.Add upd
			End If
		Next
	Next

	' Populate list with Updates to hide, while making sure
	' there are no duplicates
	if tempUpdToHide.Count > 1 Then
		Set updatesToHide = removeDuplicates(tempUpdToHide)	
	End If
	
	printf MSG_UPDATES_TO_HIDE
	
	For Each uth in updatesToHide
		PrintUpdateDetails(uth)
	Next
	
	printf MSG_SEPARATOR_LINE
End Function


' scan <notInstUpds> and hide matching update with the KB provided as argument
Sub PerformHideUpdates
	Dim upd, kbID
	
	On Error Resume Next
	For Each upd in updatesToHide
		kbID = GetKBID(upd)
		upd.isHidden = True
		printf MSG_HIDE_UPD_SUCCESS & kbID
	Next
	On Error Goto 0
	printf vbNewLine
	
End Sub


' ##########################################################################################
' ##################       FUNCTIONS CALLING (pgogram execution)       #####################
' ##########################################################################################

	printf MSG_STARTING_PROGRAM

	' read the text file that contains blacklisted updates
	Call ParseTxtFile(filename)
	
	PressKeyToContinue(MSG_START_WU_SEARCH)
	
' starts stopwatch
StartTime = Timer
printf MSG_CHRONO_START & FormatDateTime(Now,3)
	
	' searches for updates online and
	' sort the updates in Installed and Not Installed
	Set wuResult = PerformUpdateCheck
	Call ParseUpdates(wuResult)
	
ElapsedTime = Timer - StartTime
printf  MSG_CHRONO_END & ElapsedTime & " seconds." & vbNewLine
printf MSG_SEPARATOR_LINE

	PressKeyToContinue(MSG_SHOW_INST_UPDS)
	
	' Lists blacklisted Updates that are already installed
	Call ShowInstalledUpds
	
	PressKeyToContinue(MSG_SHOW_UPDS_TO_HIDE)
	
	' Lists blacklisted that are about to be hidden
	Call ShowUpdatesToHide
	
	PressKeyToContinue(MSG_HIDE_UPDATES_INFO)
	
	' Hide Updates
	Call PerformHideUpdates
	
	PressKeyToContinue(MSG_EXIT_PROGRAM)





