Imports SAPScripts
Imports Common_Functions

Module Module1


    Sub Main()

        'Dim SAPApp = CreateObject("Sapgui.ScriptingCtrl.1")
        'Dim Connection = SAPApp.OpenConnectionByConnectionString("/R/N6P/G/SPACE/M/N6P.na.pg.com")
        'Dim Session = Connection.Children(0)

        Dim Session As New SAPGUI("L6P", "AR4041", "hmetal25")

        Dim b = Session.LoggedIn
        Session.Close()
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
