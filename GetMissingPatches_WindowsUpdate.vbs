Set objFSO=CreateObject("Scripting.FileSystemObject")

' How to write file
outFile="GetMissingPatchesLogs.txt"
Set objFile = objFSO.CreateTextFile(outFile,True)

timeStart = time()
Set updateSession = CreateObject("Microsoft.Update.Session")
updateSession.ClientApplicationID = "MSDN Sample Script"

Set updateSearcher = updateSession.CreateUpdateSearcher()
WScript.Echo timeStart & " : Start search for updates..."
objFile.Write timeStart & " : Start search for updates..." & vbCrLf
Set searchResult = _
updateSearcher.Search("(IsInstalled=0 AND IsHidden=1) OR (IsInstalled=0 AND IsHidden=0) OR (IsInstalled=0 AND DeploymentAction=*)")

timeEnd = time()
timeExcute = DateDiff("s", timeStart, timeEnd)
WScript.Echo timeEnd & " : End search for updates..."
objFile.Write timeEnd & " : End search for updates..." & vbCrLf
WScript.Echo timeExcute & "s to search for updates "
objFile.Write timeExcute & "s to search for updates" & vbCrLf

If searchResult.Updates.Count = 0 Then
    WScript.Echo "There are no applicable updates. Quit"
	objFile.Write "There are no applicable updates. Quit" & vbCrLf
	objFile.Close
	WScript.Quit
End If


WScript.Echo "List of applicable items on the machine: " & searchResult.Updates.Count  
objFile.Write "List of applicable items on the machine: " & searchResult.Updates.Count & vbCRLF


For I = 0 To searchResult.Updates.Count-1
    Set update = searchResult.Updates.Item(I)
    ' WScript.Echo I + 1 & "> " & update.Title
	objFile.Write update.Title & vbCrLf
Next

WScript.Echo "Quit" & vbCRLF
objFile.Write "Quit" & vbCrLf
objFile.Close
WScript.Quit