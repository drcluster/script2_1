book = WScript.Arguments(0)
max_words = WScript.Arguments(1)

set booktext = CreateObject("Scripting.FileSystemObject").OpenTextFile(book, 1)
set frequency = CreateObject("Scripting.Dictionary")

do until booktext.AtEndOfStream
    currentLine = LCase(booktext.ReadLine())
    wordsArray = Split(currentLine, " ")
    for each item In wordsArray
		punctuationMarks = Array(",", ";", ":", "(", ")", "!", "?", ".", """", "--", "&", "$", "  ", "_")
		For i = 0 To UBound(punctuationMarks)
			item = Replace(item, punctuationMarks(i), "")
		Next
		
		Select Case item
			Case "this", "that", "these", "those" : item = "the"
			Case "is", "are", "am", "was", "were" : item = "be"
			Case "toward", "unto", "until", "till" : item = "to"
			Case "off", "off of" : item = "of"
			Case "also", "plus", "as well as" : item = "and"
			Case "an" : item = "a"
			Case "inside", "within", "into" : item = "in"
			Case "has", "had" : item = "have"
			Case "its" : item = "it"
			Case "don't", "doesn't", "didn't" : item = "not"
			Case "him", "his" : item = "he"
			Case "those", "this" : item = "that"
			Case "your", "yours" : item = "you"
			Case "does", "did" : item = "do"
			Case "is", "am", "was", "were" : item = "are"
			Case "that" : item = "this"
			Case "however", "nevertheless" : item = "but"
			Case "upon", "onto" : item = "on"
			Case "along with", "together with" : item = "with"
			Case "me", "my", "mine" : item = "i"
			Case "around", "about" : item = "at"
			Case "beside", "near", "next to" : item = "by"
			Case "them", "their", "theirs" : item = "they"
			Case "us", "our", "ours" : item = "we"
			Case "said", "says" : item = "say"
			Case "hers", "her" : item = "she"
			Case "four", "fore", "before" : item = "for"
			Case "great", "excellent", "positive", "beneficial" : item = "good"
			Case "moment", "occasion", "event" : item = "time"
			Case "job", "employment", "task", "labor" : item = "work"
			Case Else :
		End Select

        if frequency.Exists(item) then
            frequency(item) = frequency(item) + 1
        else
            frequency.Add item, 1
        end if
    next
loop


wscript.echo "CHECKING THE ZIPF's LAW"
wscript.echo "The first column is the number of corresponding words in the text and the second column is the number of words which should occur in the text according to the Zipf's law."
wscript.echo "The most popular words in " + book + " are:"

For i = 0 To max_words
    maxWord = "" : maxCount = 0
    For Each amount In frequency
        If frequency(amount) > maxCount Then
            maxWord = amount
            maxCount = frequency(amount)
        End If
    Next

    If maxCount <> 0 Then
        frequency.Remove(maxWord)
        If Not (maxWord = vbLf Or maxWord = vbCr Or maxWord = "" Or maxWord = " ") Then
            zipfLawCount = maxCount
            WScript.Echo maxWord, vbTab, maxCount, vbTab, Round(zipfLawCount / i)
        End If
    End If
Next

wscript.echo
wscript.echo "The most popular still remaining short forms in " + book + " are:"

for i = 0 to max_words
    maxWord = "" : maxCount = 0
    for each amount In frequency
        if frequency(amount) > maxCount And InStr(amount, "'") then
            maxWord = amount
            maxCount = frequency(amount)
        end if
    next

    if maxCount <> 0 then
        frequency.Remove(maxWord)
        if Not (maxWord = vbLf Or maxWord = vbCr Or maxWord = "" Or maxWord = " ") Then
            WScript.Echo maxWord, vbTab, maxCount
        end if
    end if
next