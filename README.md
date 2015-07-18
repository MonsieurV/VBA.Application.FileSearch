# Application.FileSearch replacement

`YtoFileSearch` is a VBA class replacement for the `Application.FileSearch` object that was deprecated with Office 2007.

This class aims to be used as the original [Application.FileSearch Object](https://msdn.microsoft.com/en-us/library/office/aa219847(v=office.11).aspx).

To use it, copy-paste the source code in `YtoFileSearch.vba` to a class module named `YtoFileSearch`.

Then use the `YtoFileSearch` (almost) as you used the `Application.FileSearch` object:

```
Dim fs As YtoFileSearch
Set fs = New YtoFileSearch
With fs
    .NewSearch
    .LookIn = "D:\User\Downloads\"
    .fileName = "*.pdf"
    If .Execute() > 0 Then
        Debug.Print "Found these PDF files:"
        For i = 1 To .FoundFiles.Count
           Debug.Print .FoundFiles(i)
       Next
    Else
        Debug.Print "Nothing found"
    End If
End With
```

Please open an issue for any request, the class still need improvement.

May your VBA duty be gentle, may it be short. 
