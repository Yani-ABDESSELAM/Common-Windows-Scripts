'Script to delete the Windows hidden uninstall patches older than "X" number of days
    On Error Resume NExt 
    SET oFSO = WScript.CreateObject("Scripting.FileSystemObject")
    SET oFolder = oFSO.GetFolder("C:\Windows\Installer\$PatchCache$\Managed")

    days = 300  'Enter days count you want. I'll detele older then your entered day
    For Each folder In oFolder.SubFolders
          If folder.Attributes And 2 Then
            If folder.DateLastModified < dateadd("d", -days, Now) then         

                WScript.Echo folder.Name & " last modified at " & folder.DateLastModified
                path = folder.Path
                oFSO.DeleteFolder(folder),True
                WScript.Echo path & " was hidden and deleted just now"
            Else 
                WScript.Echo folder.Name & " last modified at " & folder.DateLastModified & " newer than " & days & " days"
            End If

      End If
    Next

    If err.number<>0 then
        WScript.Echo "Script Check Failed"
        Wscript.Quit 1001
    Else 
        WScript.Echo "Script Check Passed"
        Wscript.Quit 0
    End If