Sub EEWordCount()
Dim Tbl As Table, d As Long, t As Long
Dim p As Paragraph, w As Long

With ActiveDocument
  For Each Tbl In .Tables
    With Tbl
      t = t + .Range.ComputeStatistics(wdStatisticWords)
    End With
  Next
  
  For Each p In .Paragraphs
    With p
      If .Style = "Caption" Then
        w = w + .Range.ComputeStatistics(wdStatisticWords)
      End If
    End With
  Next
  
  d = .Range.ComputeStatistics(wdStatisticWords)
End With



  
MsgBox "There are:" & vbCr & _
  d & " words in the document body, including" & vbCr & _
  t & " words in tables." & vbCr & _
  w & " words as Captions." & vbCr & vbCr & "There are" & vbCr & _
  d - t - w & " words in the document body, excluding tables, captions and footnotes."
End Sub

