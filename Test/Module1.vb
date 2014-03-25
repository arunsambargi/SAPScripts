Imports SAPScripts
Imports Common_Functions
Imports OfficeOpenXml

Module Module1


    Sub Main()

        Dim S As New Credits_ST100_Scripting("L6P", "CA5482", "cezane4")
        Dim I As New List(Of String)
        Dim C As New List(Of String)
        I.Add("4610004514")
        I.Add("6540152539")
        C.Add("4610005514")
        S.Execute(1, I, "0015000877", "501", "", C, "4610004514")

    End Sub

    Sub Nombres()

        Dim Session As New SAPGUI("L7P")
        Dim TP As New ExcelPackage(New IO.FileInfo("C:\Temp\Book1.xlsx"))
        Dim WS As ExcelWorksheet = TP.Workbook.Worksheets(1)
        Dim I As Integer = 1
        Do While WS.Cells("A" & I).Value <> ""

            Session.StartTransaction("SU01D")
            Session.FindByNameEx("USR02-BNAME", 32).Text = WS.Cells("A" & I).Value
            Session.FindById("wnd[0]/tbar[1]/btn[7]").press()
            If Session.FindById("wnd[1]") Is Nothing Then
                WS.Cells("B" & I).Value = Session.FindByNameEx("ADDR3_DATA-NAME_TEXT", 31).Text
            Else
                WS.Cells("B" & I).Value = "The user does not exist"
            End If

            I += 1

        Loop

        TP.Save()

    End Sub

End Module
