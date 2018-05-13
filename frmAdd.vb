    Private Sub insertVal(ByVal Table As String, ByVal strValue As String)
        found = False
        SQLa = "SELECT * FROM " & Table

        CallAccess()
        RSa.MoveFirst()

        Do Until (RSa.EOF)
            If RSa.Fields(Table).Value = strValue Then
                found = True
                Exit Do
            Else
                RSa.MoveNext()
            End If
        Loop
        If found <> True Then
            RSa.MoveLast()
            RSa.AddNew()
            RSa.Fields(Table).Value = strValue
        End If
        RSa.Update()
        CloseAccess()

    End Sub
End Class
