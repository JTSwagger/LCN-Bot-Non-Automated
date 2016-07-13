Module Module2
    Function ReturnMostCommonstring(ByVal StringList As List(Of String)) As String()()
        Dim spot As Integer = 0
        Dim word As String = ""
        Dim StrCompare As String = ""
        Dim place_holder As Integer = 0
        Dim strStart As Integer = 0
        Dim strEnd As Integer = 1
        Dim subStart As Integer = 0
        Dim subEnd As Integer = 0
        Dim NumOccur(0)() As String

        For Each word In StringList
            StrCompare = word.Substring(strStart, strEnd)
            For X As Integer = 0 To StringList.Count
                If spot <> X Then
                    If StringList.Item(X).Contains(StrCompare) Then
                        Console.WriteLine(StrCompare & " exists in " & word)
                    End If
                End If
            Next
        Next
        Return NumOccur
    End Function
End Module
