Imports SAPScripts

Module Module1

    Sub Main()

        Dim Session As New SAPGUI("N6P")
        Session.DisplayPO("4503401042")

        'Dim C As New List(Of String)
        'Dim I As New List(Of String)

        'C.Add("4610000478")
        'I.Add("6540034885")

        'Dim S As New Credits_ST100_Scripting("A6P", "CA5482", "suntzu1")
        'S.Execute(1, I, "0075601050", "604", "", C)

    End Sub


End Module
