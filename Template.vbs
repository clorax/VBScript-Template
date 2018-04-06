Option Explicit

'=============================================================================================================
Const CScriptVersion 	= "0.0.1  (dd/mm/yyyy)"
Const CScriptName	= "Script Name"
Const CScriptOwner	= "Your Name"
'=============================================================================================================

'Objects
Dim WshShell         : Set WshShell = CreateObject ("WScript.Shell")
Dim WshSysEnv        : Set WshSysEnv = WshShell.Environment("Process")
Dim FileSystemObject : Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
Dim WshNetwork       : Set WshNetwork = CreateObject("WScript.Network")
Dim objWMIService    : Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")


'Current PATH
Dim GpathScript : GpathScript = replace(Wscript.ScriptFullName,Wscript.ScriptName,"")

'Indentation for TraceDebug
Dim GIndent
Dim GFunctionTime (100)
Dim GdebugFile, GdebugCentralFile
Dim generalError : generalError = 0
Dim GLogMsg


'Constants
Const CServerLog = "" 	'The path to the log server
Const CDebug = 1	'Cumulative Flag:  0 = No Debug ; 1 = local file log ; 2 = central server log ; 8 = Screen log ; 16 = Event Log ; 32 = Mail

Const CSmartHost = ""	'SMTP Gateway
Const CMailDestin = "" 	'Mail Destin
Const CMailCC = "" 	'Mail CC
Const CSendUsername = "" 'SMTP Username
Const CSendPassword = "" 'SMTP Password
Const CSMTPport = 25	'SMTP Port
Const CSMTPAuth = 0	'0 = Anonymous, 1 = Basic, 2 = NTLM
Const CSMTPuseSSL = 0	'0 = no, 1 = yes
Dim CMailFrom : CMailFrom = CScriptName & "@"	'Mail From

Const CEventSuccess = 96	'Success Event ID
Const CEventError = 97	'Failure Event ID


'use the US date and time
setlocale("en-us")

Main

Function Main()
	Main = 0
	
	TraceDebug ""
	TraceDebug "  ************************************************************"
	TraceDebug "              BEGIN "
	TraceDebug ""
	TraceDebug "  Script information : "
	TraceDebug "     Name    : " & CScriptName
	TraceDebug "     Version : " & CScriptVersion
	TraceDebug "     Owner   : " & CScriptOwner
	TraceDebug ""
	StartFunction "Main"
	









	If (CDebug AND 16) Then UpdateEventLog	
	If (CDebug AND 32) Then 
		If generalError = 0 Then
			SendMail "Success", "Success"
		Else
			SendMail "Failed", "Failed"
		End If
	End If

	Main = generalError
	
	EndFunction "Main : " & Main

	TraceDebug ""
	TraceDebug "              END"
	TraceDebug "  ************************************************************"
	
	Wscript.quit Main

End Function


'
' Function 	: TraceDebug
' Description 	: Write the log file
' Parameter 	: Msg (R)	: the message to write in the log file
' Result 	: none
'----------------------------------------------------------------------------------------

Function TraceDebug (byval Msg)

	if left(Msg,1) <> "*" then 
		Msg = "  . " & Msg
	Else
		Msg = " " & Msg
	End If
	
	If GIndent<0 Then GIndent = 0
	Dim LogMsg : LogMsg = string(2*GIndent," ") & Msg
	

	if (CDebug AND 1) Then LocalLog LogMsg
	if (CDebug AND 2) Then ServerLog LogMsg
	if (CDebug AND 8) Then Wscript.Echo LogMsg	

	GLogMsg = GLogMsg & vbNewLine & Now & " - " & WshSysEnv("COMPUTERNAME") & " - " & WshSysEnv("USERNAME") & " - " & LogMsg

End Function


'
' Function 	: LocalLog
' Description 	: Write the local log file
' Parameter 	: Msg (R)	: the message to write in the log file
' Result 	: none
'----------------------------------------------------------------------------------------
Function LocalLog (byval Msg)

	On Error Resume Next
		GdebugFile.WriteLine Now & " - " & WshSysEnv("COMPUTERNAME") & " - " & WshSysEnv("USERNAME") & " - " & Msg
		If Err.Number <> 0 Then
			GdebugFile.Close
			Set GdebugFile = nothing
			Set GdebugFile = FileSystemObject.OpenTextFile(GpathScript & GenerateDate & "-" & CScriptName & ".log",8,True)
			GdebugFile.WriteLine Now & " - " & WshSysEnv("COMPUTERNAME") & " - " & WshSysEnv("USERNAME") & " - " & Msg
		End If
	On Error Goto 0

End Function


'
' Function 	: ServerLog
' Description 	: Write the central server log file
' Parameter 	: Msg (R)	: the message to write in the log file
' Result 	: none
'----------------------------------------------------------------------------------------
Function ServerLog (byval Msg)

	On Error Resume Next
		GdebugCentralFile.WriteLine Now & " - " & WshSysEnv("COMPUTERNAME") & " - " & WshSysEnv("USERNAME") & " - " & Msg
		If Err.Number <> 0 Then
			GdebugCentralFile.Close
			Set GdebugCentralFile = nothing
			Set GdebugCentralFile = FileSystemObject.OpenTextFile(CServerLog & "\" & CScriptName & "\online_" & WshSysEnv("COMPUTERNAME") & "_" & WshSysEnv("USERNAME") & ".log",8,True)
			GdebugCentralFile.WriteLine Now & " - " & WshSysEnv("COMPUTERNAME") & " - " & WshSysEnv("USERNAME") & " - " & Msg
		End If
	On Error Goto 0

End Function


'
' Function 	: UpdateEventLog
' Description 	: Write the Event Log
' Parameter 	: none
' Result 	: none
'----------------------------------------------------------------------------------------
Function UpdateventLog ()

	If generalError = 0 Then
		CreateEvent "INFORMATION", CScriptName & " ended successfully", CEventSuccess
	Else
		CreateEvent "ERROR", CScriptName & " ended with errors", CEventError
	End If

End Function


'
' Function 	: SendMail
' Description 	: Send an EMail
' Parameter 	: MailSubject (R)	: Subject
'              	: MailText (R)	: Body
' Result 		: None
'----------------------------------------------------------------------------------------
Sub SendMail (MailSubject, MailText)
	
	StartFunction "SendMail"
		
		Dim iMsg
		Dim iConf
		Dim Flds
		Dim Latestlog
		Const cdoSendUsingPort = 2
		
		On Error Resume Next
			'Create the message object.
			Set iMsg = CreateObject("CDO.Message")
			
			'Create the configuration object.
			Set iConf = iMsg.Configuration
			
			'Set the fields of the configuration object
			With iConf.Fields
				.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
				.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = CSmartHost
				.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CSMTPport
				.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = CSMTPAuth
				.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = CSendUsername
				.item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = CSendPassword
				.item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = CSMTPuseSSL
				.Update
			End With
			
			'Set the To, From, Subject, and Body properties of the message.
			With iMsg
		
		
				.To = CMailDestin
				.Cc = CMailCC 
				.From = CMailFrom
				.Subject = "[" & CScriptName & "] " & WshSysEnv("COMPUTERNAME") & " - " & MailSubject
				.TextBody = MailText
		'		.HTMLBody = MailText
		
				GdebugFile.Close
				If FileSystemObject.FileExists(GpathScript & GenerateDate & "-" & CScriptName & ".log") Then
					.AddAttachment GpathScript & GenerateDate & "-" & CScriptName & ".log"
				End If 			
				.Send
			End With
			set iMsg = Nothing
			ManageError "Try to send mail"
		On Error GoTo 0
		
		
	EndFunction "SendMail"
	
End Sub

'
' Function 	: CreateEvent
' Description 	: Create an event in the Event Log
' Parameter 	: strType	: Type of message (Information, Warning, etc)
'				: strMsg	: Event text
'				: iID		: Event ID number
' Result 	: none
'----------------------------------------------------------------------------------------
Sub CreateEvent(strType, strMsg, iID)
	StartFunction "CreateEvent"
		on error resume next
			TraceDebug ("eventcreate /T " & strTYPE & " /ID " & iID & " /L APPLICATION /SO " & chr(34) _
				& CScriptName & chr(34) & " /D " & chr(34) & strMsg & chr(34))
			WshShell.run ("eventcreate /T " & strTYPE & " /ID " & iID & " /L APPLICATION /SO " & chr(34) _
				& CScriptName & chr(34) & " /D " & chr(34) & strMsg & chr(34))
		on error goto 0
	EndFunction "CreateEvent"
End Sub

'
' Function 	: ManageError
' Description 	: Write an error in the log file and return the error number
' Parameter 	: Msg (R)	: The message to write in case of error
' Result 	: the error number (0 if it's OK)
'----------------------------------------------------------------------------------------
Function ManageError(Msg)
        ManageError = Err.number
        if Err.Number <> 0 then
                TraceDebug "$$ Error $$ [" & Err.Number & "] - [" & Err.Description & "] - [" & Msg & "]"
				generalError = 1
                Err.Clear
        else
               ' TraceDebug "No error " & Msg
        end if
end Function

'
' Function 	: StartFunction
' Description 	: Write a line into the log file with '* Begin <Msg>' and indent the next message
' Parameter 	: Msg (R)	: the name of the function
' Result 	: None
'----------------------------------------------------------------------------------------
Function StartFunction (Msg)
	
	GIndent = GIndent + 1
	If GIndent < UBound(GFunctionTime) Then GFunctionTime(GIndent) = Now
	TraceDebug "* Begin " & Msg
End function


'
' Function 	: EndFunction
' Description 	: Write a line into the log file with '* End <Msg>' and indent the next message
' Parameter 	: Msg (R)	: the name of the function
' Result 	: None
'----------------------------------------------------------------------------------------
Function EndFunction (Msg)
	
	If GIndent < UBound(GFunctionTime) Then Msg = Msg & " - in " & DateDiff("s",GFunctionTime(GIndent), Now) & " (s) - " & DateDiff("n",GFunctionTime(GIndent), Now) & " (mn)"
	TraceDebug "* End " & Msg
	GIndent = GIndent - 1
	
End Function


'
' Function 	: 
' Description 	: 
' Parameter 	: Msg (R)	: 
' Result 	: 
'----------------------------------------------------------------------------------------
Function GenerateDate ()

	GenerateDate = DatePart("yyyy",Date) _
    	& Right("0" & DatePart("m",Date), 2) _
    	& Right("0" & DatePart("d",Date), 2) 

End Function