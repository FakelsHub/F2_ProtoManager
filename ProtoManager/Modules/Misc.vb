Module Misc

    ''' <summary>
    ''' Удаляет пустые строки в конце массива и лишние пробелы в строке.
    ''' </summary>
    Friend Function ClearEmptyLines(ByRef lst As String()) As String()
        Dim count As Integer = UBound(lst)
        Dim n As Integer
        For n = 0 To count
            lst(n) = lst(n).Trim
        Next
        For n = count To 0 Step -1
            If (lst(n).Length > 0) Then
                Exit For
            End If
        Next
        If (count <> n) Then
            ReDim Preserve lst(n)
        End If
        Return lst
    End Function

End Module
