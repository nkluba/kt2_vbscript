Option Explicit

Dim objFSO, objTextFile
Dim strText, arrWords, word, dict
Dim i, intUserSpecifiedNum, intTotalWords
Dim arrKeys, arrItems

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objTextFile = objFSO.OpenTextFile("C:\Users\luban\Downloads\sample.txt", 1)

strText = objTextFile.ReadAll

objTextFile.Close

arrWords = Split(strText)

Set dict = CreateObject("Scripting.Dictionary")

For Each word In arrWords
    If dict.Exists(word) Then
        dict.Item(word) = dict.Item(word) + 1
    Else
        dict.Add word, 1
    End If
Next

intTotalWords = UBound(arrWords) + 1

intUserSpecifiedNum = 10

arrKeys = dict.Keys
arrItems = dict.Items

Call BubbleSort(arrKeys, arrItems)

For i = 0 To intUserSpecifiedNum - 1
    WScript.Echo "Word: " & arrKeys(i) & ", Actual Occurrences: " & arrItems(i) & ", Predicted Occurrences: " & intTotalWords / (i + 1)
Next

Sub BubbleSort(arr1, arr2)
    Dim i, j, temp1, temp2
    For i = UBound(arr1) - 1 To 0 Step -1
        For j= 0 To i
            If arr2(j) < arr2(j + 1) Then
                temp1 = arr1(j)
                temp2 = arr2(j)
                arr1(j) = arr1(j + 1)
                arr2(j) = arr2(j + 1)
                arr1(j + 1) = temp1
                arr2(j + 1) = temp2
            End If
        Next
    Next
End Sub
