Imports SAPScripts
Imports OfficeOpenXml
Imports Common_Functions

Module Module1

    Dim SQL As New SQL_Server("MXL0221QY0\SQLEXPRESS", "developer", "procter", "PSSD_LBI")
    Dim MF As New MyFunctions_Class

    Sub Main()

        Dim Session As New SAPGUI("L7P", "CA5482", "control1", "control2")


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
