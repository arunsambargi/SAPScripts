Imports SAPScripts
Imports OfficeOpenXml
Imports Common_Functions

Module Module1

    Dim SQL As New SQL_Server("MXL0221QY0\SQLEXPRESS", "developer", "procter", "PSSD_LBI")
    Dim MF As New MyFunctions_Class

    Sub Main()

        Dim Session As New SAPGUI("ANP", "CA5482", "tsuntzu1", , "430")

        'Dim C As New List(Of String)
        'Dim I As New List(Of String)

        'C.Add("4610000478")
        'I.Add("6540034885")

        'Dim S As New Credits_ST100_Scripting("A6P", "CA5482", "suntzu1")
        'S.Execute(1, I, "0075601050", "604", "", C)

    End Sub


End Module
