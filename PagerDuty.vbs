
' This script sends the content of text files in the QueuePath to the URL.
'
' Text files should contain all of the JSON neccessary to trigger, acknowledge, or
' resolve an incident in PagerDuty as defined in the PagerDuty Developer documentation:
' https://developer.pagerduty.com/documentation/integration/events

On Error Resume Next

' Constants and shell object used for logging to the Windows Application Event Log

Const EVENT_SUCCESS	= 0
Const EVENT_ERROR 	= 1
Const EVENT_WARNING = 2
Const EVENT_INFO 	= 4

Set objShell = Wscript.CreateObject("WScript.Shell")

' Set the API endpoint and alert file directory we're working with, and whether
' you want successful calls to be logged in the Windows Application Event Log.
' NOTE: Trailing backslash is required for QueuePath!

Dim URL, QueuePath, LogSuccess

URL = "https://events.pagerduty.com/generic/2010-04-15/create_event.json"
QueuePath = "C:\PagerDuty\Queue\"
LogSuccess = True

' Get list of filenames in the queue directory

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(QueuePath)
Set colFiles = objFolder.Files

For Each objFile in colFiles

	' Set filename variables and check for the extension ".txt" so we don't mess
	' with lock files or anything else in the queue that isn't expected to be there.

	Dim AlertFileName, AlertFileExtension, AlertFile

	AlertFileName = objFile.Name
	AlertFileExtension = objFSO.GetExtensionName(AlertFileName)
	AlertFile = QueuePath & AlertFileName

	If AlertFileExtension = "txt" Then

		' Check for a lock and quit if we're already working on this alert in another
		' process, or create a lock if we're going to start working on this alert.

		Dim PostBody, AlertLockFile, AlertFileContent, Status, Response

		AlertLockFile = AlertFile & ".lock"

		If objFSO.FileExists(AlertLockFile) Then
			WScript.Echo "Lock file already exists: " & AlertLockFile & ". Moving on to next alert file."
		Else

			WScript.Echo "Creating lock file: " & AlertLockFile
			objFSO.CreateTextFile(AlertLockFile)

			' Open and get the alert file content, escaping backslashes with an additional backslash

			Err.Clear
			Set AlertFileContent = objFSO.OpenTextFile(AlertFile,1)
			PostBody = AlertFileContent.ReadAll
			PostBody = Replace(PostBody, "\", "\\")
			AlertFileContent.Close

			If Err.Number <> 0 Then
				WScript.Echo "ERROR: Couldn't read alert file: " & AlertFile & vbNewLine &_
					"Check Windows Application Event Log for details."

				objShell.LogEvent EVENT_ERROR, "Couldn't read alert file." & vbNewLine & vbNewLine &_
					"File name: " & AlertFile & vbNewLine &_
					"Error Number: " & Err.Number & vbNewLine &_
					"Source: " & Err.Source & vbNewLine &_
					"Description: " & Err.Description
			Else
				' Send the alert file content to PagerDuty and check response

				Set objHTTP = CreateObject("MSXML2.XMLHTTP.3.0")

				Err.Clear
				WScript.Echo "Body being sent to PagerDuty is: " & PostBody
				objHTTP.Open "POST", URL, False
				objHTTP.Send PostBody

				' Log connection failures in case the system isn't able to resolve
				' the domain or server isn't responding at all.

				If Err.Number <> 0 Then
					WScript.Echo "ERROR: Couldn't connect or send data to PagerDuty. Check Windows Application Event Log for details."

					objShell.LogEvent EVENT_ERROR, "Couldn't connect or send data to PagerDuty." & vbNewLine & vbNewLine &_
						"Error Number: " & Err.Number & vbNewLine &_
						"Source: " & Err.Source & vbNewLine &_
						"Description: " & Err.Description
				Else
					Status = objHTTP.Status
					Response = objHTTP.responseText
					WScript.Echo "Response from PagerDuty is: [" & Status & "] " & Response

					' Remove the alert from the queue once it has been accepted by PagerDuty,
					' or log why the event wasn't accepted by PagerDuty.

					If Status = 200 Then
						WScript.Echo "Deleting alert file: " & AlertFile

						objFSO.DeleteFile(AlertFile)

						If LogSuccess = True Then
							objShell.LogEvent EVENT_SUCCESS, "PagerDuty accepted event with data:" & vbNewLine & vbNewLine &_
								PostBody & vbNewLine & vbNewLine &_
								"Response was:" & vbNewLine & vbNewLine &_
								"[" & Status & "] " & Response
						End If
					Else
						WScript.Echo "Non-200 response received. Keeping alert file in queue: " & AlertFile

						objShell.LogEvent EVENT_ERROR, "PagerDuty did not accept event with data:" & vbNewLine & vbNewLine &_
							PostBody & vbNewLine & vbNewLine &_
							"Response was:" & vbNewLine & vbNewLine &_
							"[" & Status & "] " & Response
					End If
				End If
			End If

			' Remove the lock file. This will let us try again in case PagerDuty couldn't be reached.

			WScript.Echo "Deleting lock file: " & AlertLockFile
			objFSO.DeleteFile(AlertLockFile)
		End If
	End If
Next