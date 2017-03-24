Attribute VB_Name = "modWord"
Dim oWord As Word.Application
Dim oDoc As Word.Document

Public Sub StartWord()

Set oWord = New Word.Application

End Sub

Public Sub ExitWord()

oWord.Visible
oWord.Quit (wdDoNotSaveChanges)
Set oWord = Nothing

End Sub

Public Sub OpenDoc(sFileName As String)

Set oDoc = oWord.Documents.Open(FileName)

End Sub

Public Sub CloseDoc()

oDoc.Close (wdDoNotSaveChanges)

End Sub

Public Sub PrintDoc()

Dim n As Integer
Dim TVar As Variant
Dim Pages As Integer
Pages = oDoc.ComputeStatistics(wdStatisticPages)
For n = 1 To Pages

TVar = Pages - (n - 1)
DoEvents
Call oDoc.PrintOut(, , wdPrintRangeOfPages, , , , , , TVar)

Next n

End Sub
