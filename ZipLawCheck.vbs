Option Explicit

Sub CheckZipfsLaw(filePath, intUserSpecifiedNum)
    Dim objFSO, objTextFile
    Dim strText, arrWords, word, dict, dictApostrophe
    Dim i, intTotalWords, normalizedWord
    Dim arrKeys, arrItems
    Dim wordForms
	Dim regex

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTextFile = objFSO.OpenTextFile(filePath, 1)

    strText = LCase(objTextFile.ReadAll)
    objTextFile.Close
	
	strText = Replace(Replace(strText, Chr(13), " "), Chr(10), " ")
	Set regex = New RegExp
	regex.Pattern = "[^a-zA-Z\s']"
	regex.Global = True
	strText = regex.Replace(strText, "")

    Set wordForms = CreateObject("Scripting.Dictionary")
    wordForms.Add "be", Array("is", "are", "am", "was", "were")
	wordForms.Add "he", Array("him", "his")
    wordForms.Add "a", Array("an")
    wordForms.Add "have", Array("has", "had")
    wordForms.Add "it", Array("its")
    wordForms.Add "not", Array("don't")
    wordForms.Add "you", Array("your")
    wordForms.Add "do", Array("does", "did")
    wordForms.Add "they", Array("their")
    wordForms.Add "we", Array("our")
    wordForms.Add "say", Array("says", "said")

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
	Set dictApostrophe = CreateObject("Scripting.Dictionary")

	For Each word In arrWords
		If word <> "" Then
			If InStr(word, "'") > 0 Then
				If dictApostrophe.Exists(word) Then
					dictApostrophe.Item(word) = dictApostrophe.Item(word) + 1
				Else
					dictApostrophe.Add word, 1
				End If
			Else
				If dict.Exists(word) Then
					dict.Item(word) = dict.Item(word) + 1
				Else
					dict.Add word, 1
				End If
			End If
		End If
	Next

    intTotalWords = UBound(arrWords) + 1

    WScript.Echo "The most popular words in " & filePath & " are: " & vbCrLf
	arrKeys = dict.Keys
	arrItems = dict.Items
	Call BubbleSort(arrKeys, arrItems)
	For i = 0 To intUserSpecifiedNum - 1
		WScript.Echo arrKeys(i) & " " & arrItems(i) & " " & intTotalWords / (i + 1)
	Next

	WScript.Echo vbCrLf & "The most popular still remaining short forms in " & filePath & " are: " & vbCrLf
	arrKeys = dictApostrophe.Keys
	arrItems = dictApostrophe.Items
	Call BubbleSort(arrKeys, arrItems)
	For i = 0 To intUserSpecifiedNum - 1
		WScript.Echo arrKeys(i) & " " & arrItems(i) & " " & intTotalWords / (i + 1)
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
