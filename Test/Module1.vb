Imports SAPScripts
Imports Common_Functions

Module Module1


    Sub Main()

        Dim S As New Credits_ST100_Scripting("N6P", "CA5482", "cezane3")
        Dim I As New List(Of String)
        Dim C As New List(Of String)
        I.Add("7780036602")
        C.Add("4110000124")

        S.Execute(1, I, "0015256391", "013", "", C)


    End Sub



End Module
