<div align="center">

## Embed Binary Files In Programs


</div>

### Description

Embed binary files within your programs to be written out to the file system once the program runs. This demonstration creates a Word document.
 
### More Info
 
Assumes you know how to get the hex read-out of binary files. If not, many hex editors are available on PSCode.com


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lewis E\. Moten III](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lewis-e-moten-iii.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lewis-e-moten-iii-embed-binary-files-in-programs__1-35007/archive/master.zip)





### Source Code

```
' Example of use:
' Call CreateFile("c:\test.doc")
Sub CreateFile(ByRef pstrFileName)
 ' Creates a binary file in the location provided.
 ' This code creates a blank word document.
 ' However, it can be modified to create a different
 ' file of your choice.
 Dim llngIndex As Long
 Dim lbytTransferAry(0) As Byte
 Dim llngFileNum As Long
 Dim llngMaxIndex As Long
 Dim lstrData As String
 Dim lstrByte As String
 Dim llngPosition As Long
 ' The following is a hex readout of a new word document
 ' Replace this code with the contents of a seperate file
 ' that you wish to create
 lstrData = _
 "7B 5C 72 74 66 31 5C 61 6E 73 69 5C 61 6E 73 69" & _
 "63 70 67 31 32 35 32 5C 64 65 66 66 30 5C 64 65" & _
 "66 6C 61 6E 67 31 30 33 33 7B 5C 66 6F 6E 74 74" & _
 "62 6C 7B 5C 66 30 5C 66 73 77 69 73 73 5C 66 63" & _
 "68 61 72 73 65 74 30 20 41 72 69 61 6C 3B 7D 7D" & _
 "0D 0A 5C 76 69 65 77 6B 69 6E 64 34 5C 75 63 31" & _
 "5C 70 61 72 64 5C 66 30 5C 66 73 32 30 5C 70 61" & _
 "72 0D 0A 7D 0D 0A 00"
 ' Remove white space (only there for readablity)
 lstrData = Replace(lstrData, " ", "")
 ' Determine max number of hex characters
 llngMaxIndex = Len(lstrData)
 ' Ignore errors
 On Error Resume Next
 ' Attempt to delete existing file (causes an error if not exists)
 FileSystem.Kill pstrFileName
 ' stop ignoring errors
 On Error GoTo 0
 ' Get a reference number to use for connecting to file
 llngFileNum = FreeFile
 ' Open file to be written
 Open pstrFileName For Binary As llngFileNum
 ' Loop through each hex byte value (byte = 2 hex characters)
 For llngIndex = 1 To llngMaxIndex Step 2
 ' Parse hex byte
 lstrByte = Mid(lstrData, llngIndex, 2)
 ' Convert data type to byte and store within array
 lbytTransferAry(0) = CByte("&H" & lstrByte)
 ' Determine where the position of the byte is in the file
 llngPosition = ((llngIndex - 1) / 2) + 1
 ' Save byte in proper position
 Put #llngFileNum, llngPosition, lbytTransferAry
 Next
 ' Close file
 Close llngFileNum
End Sub
```

