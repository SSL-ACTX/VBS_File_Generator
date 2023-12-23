' SSL-ACTX
' 2023
On Error Resume Next
Dim objFSO, currentFile, destinationFolder, numCopies, i, copyName, destinationScript, randomStream, randomData, outFile

' Create File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the current script file
currentFile = WScript.ScriptFullName

' Set the destination folder (customize this path)
destinationFolder = "destination"

' Create the destination folder if it doesn't exist
If Not objFSO.FolderExists(destinationFolder) Then
    objFSO.CreateFolder destinationFolder
End If

' Choose a random file name for the script
Randomize
Dim randomScriptName
randomScriptName = "cr_" & Int((1000 * Rnd) + 1) & "_" & Replace(WScript.ScriptName, " ", "_")
destinationScript = destinationFolder & "\" & randomScriptName

' Copy the script to the destination folder with a random file name
objFSO.CopyFile currentFile, destinationScript, True

' Number of copies to generate
numCopies = 40

' Generate random files
For i = 1 To numCopies
    copyName = "cp_" & i & "_"
    For j = 1 To 8
        copyName = copyName & Chr(Int((25 * Rnd) + 65))
    Next
    destinationFile = destinationFolder & "\" & copyName

    ' Generate random data for the files
    Randomize
    Set randomStream = CreateObject("MSXML2.DOMDocument.6.0").CreateElement("Random")
    randomStream.DataType = "bin.base64"
    ' Adjust the size here (1MB = 1024 * 1024 bytes)
    randomStream.Text = String(1024 * 1024, "A")
    randomData = Mid(randomStream.NodeTypedValue, 1, 1024 * 1024)

    ' Write random data to a file
    Set outFile = objFSO.CreateTextFile(destinationFile, True)
    outFile.Write randomData
    outFile.Close
Next

' Clean up objects
Set objFSO = Nothing
