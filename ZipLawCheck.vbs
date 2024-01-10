Option Explicit

Sub CheckZipfsLaw(filePath, intUserSpecifiedNum)
    Dim objFSO, objTextFile
    Dim strText, arrWords, word, dict
    Dim i, intTotalWords, normalizedWord
    Dim arrKeys, arrItems
    Dim wordForms

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTextFile = objFSO.OpenTextFile(filePath, 1)

    strText = objTextFile.ReadAll
    objTextFile.Close

    Set wordForms = CreateObject("Scripting.Dictionary")
    wordForms.Add "the", Array("The")
    wordForms.Add "be", Array("Be", "is", "Is", "are", "Are", "am", "Am", "was", "Was", "were", "Were")
    wordForms.Add "to", Array("To")
    wordForms.Add "of", Array("Of")
    wordForms.Add "and", Array("And")
    wordForms.Add "a", Array("A", "an", "An")
    wordForms.Add "in", Array("In")
    wordForms.Add "that", Array("That")
    wordForms.Add "have", Array("Have", "has", "Has", "had", "Had")
    wordForms.Add "i", Array("I")
    wordForms.Add "it", Array("It", "its", "Its")
    wordForms.Add "for", Array("For")
    wordForms.Add "not", Array("Not")
    wordForms.Add "on", Array("On")
    wordForms.Add "with", Array("With")
    wordForms.Add "as", Array("As")
    wordForms.Add "you", Array("You", "your", "Your")
    wordForms.Add "do", Array("Do", "does", "Does", "did", "Did")
    wordForms.Add "at", Array("At")
    wordForms.Add "this", Array("This")
    wordForms.Add "but", Array("But")
    wordForms.Add "by", Array("By")
    wordForms.Add "from", Array("From")
    wordForms.Add "they", Array("They", "their", "Their")
    wordForms.Add "we", Array("We", "our", "Our")
    wordForms.Add "say", Array("Say", "says", "Says", "said", "Said")
    wordForms.Add "or", Array("Or")

    arrWords = Split(strText)

    For i = 0 To UBound(arrWords)
        For Each word In wordForms.Keys
            For Each normalizedWord In wordForms.Item(word)
                If arrWords(i) = normalizedWord Then
                    arrWords(i) = word
                End If
            Next
        Next
    Next

    Set dict = CreateObject("Scripting.Dictionary")

    For Each word In arrWords
        If dict.Exists(word) Then
            dict.Item(word) = dict.Item(word) + 1
        Else
            dict.Add word, 1
        End If
    Next

    intTotalWords = UBound(arrWords) + 1

    arrKeys = dict.Keys
    arrItems = dict.Items

    Call BubbleSort(arrKeys, arrItems)

    For i = 0 To intUserSpecifiedNum - 1
        WScript.Echo "Word: " & arrKeys(i) & ", Actual Occurrences: " & arrItems(i) & ", Predicted Occurrences: " & intTotalWords / (i + 1)
    Next
End Sub

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

Dim filePath, intUserSpecifiedNum

If WScript.Arguments.Count >= 2 Then
    filePath = WScript.Arguments(0)
    intUserSpecifiedNum = WScript.Arguments(1)
    Call CheckZipfsLaw(filePath, intUserSpecifiedNum)
Else
    WScript.Echo "Please provide a file path and a number as arguments."
End If
