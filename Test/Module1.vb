Imports SAPScripts
Imports Common_Functions

Module Module1


    Sub Main()

        'Dim S As String = "hmetal34"
        'Dim A As String = ""
        'Dim N As String = ""
        'For Each C As Char In S.ToCharArray
        '    If IsNumeric(C) Then
        '        N = N & C
        '    Else
        '        A = A & C
        '    End If
        'Next
        'N = CStr(Val(N) + 1)
        'S = A & N

        Dim Session As New SAPGUI("N6A", "BV7795", "114116")


        Dim T As Object = Session.FindById("GRID")
        Dim I As Integer = 0
        For Each Row As Object In T.Rows
            If Row.Item(I).text = "" Then
                Exit For
            End If
            I += 1
        Next

        T.Rows(I).Selected = True
        Session.FindById("BOTON").Press()
        T.rows(I).Item(7) = "Comment"


    End Sub

End Module
