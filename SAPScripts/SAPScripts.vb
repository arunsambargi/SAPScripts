Imports System.ComponentModel
Imports System.Threading

Public Class SAPGUI

    Private SAPApp As Object = Nothing
    Private Connection As Object = Nothing
    Private Session As Object = Nothing
    Private Servers As New DataTable
    Private NPass As String = Nothing
    Private Sts As String = Nothing

    Private LI As Boolean = False

    Public Property LoggedIn() As Boolean

        Get
            LoggedIn = LI
        End Get
        Set(ByVal value As Boolean)
            LI = value
        End Set

    End Property

    Public ReadOnly Property StatusBarMessageType() As Char

        Get
            StatusBarMessageType = Session.findById("wnd[0]/sbar").MessageType
        End Get

    End Property

    Public ReadOnly Property StatusBarText() As String

        Get
            StatusBarText = Session.findById("wnd[0]/sbar").text
        End Get

    End Property

    Public ReadOnly Property SAP_Session_Obj As Object

        Get
            SAP_Session_Obj = Session
        End Get

    End Property

    Public ReadOnly Property NewPassword As String

        Get
            NewPassword = NPass
        End Get

    End Property

    Public ReadOnly Property Status As String

        Get
            Status = Sts
        End Get

    End Property

    Public Sub New(ByVal Box As String, Optional ByVal User As String = Nothing, Optional ByVal Password As String = Nothing, Optional ByRef NewPass As String = Nothing)

        Try
            Enable_GUI_Theme()
            SAPApp = CreateObject("Sapgui.ScriptingCtrl.1")
            If User Is Nothing Or Password Is Nothing Then
                Connection = SAPApp.OpenConnection(GetSSOConnString(Box), True)
            Else
                Connection = SAPApp.OpenConnectionByConnectionString(GetConnString(Box))
            End If
            Session = Connection.Children(0)
            If Not User Is Nothing And Not Password Is Nothing Then
                Session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = User
                Session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = Password
                Session.findById("wnd[0]").sendVKey(0)
            End If
            Do While Not Session.findById("wnd[1]", False) Is Nothing
                If Not Session.findById("wnd[1]/usr/pwdRSYST-NCODE", False) Is Nothing Then
                    If Not NewPass Is Nothing Then
                        Session.findById("wnd[1]/usr/pwdRSYST-NCODE").Text = NewPass
                        Session.findById("wnd[1]/usr/pwdRSYST-NCOD2").Text = NewPass
                        Session.findById("wnd[1]/tbar[0]/btn[0]").Press()
                        NPass = NewPass
                    Else
                        Session.findById("wnd[1]/tbar[0]/btn[12]").Press()
                        Exit Sub
                    End If
                End If
                If Session.ActiveWindow.Text = "SAP" Then
                    Session.findById("wnd[1]/tbar[0]/btn[12]").Press()
                    Exit Sub
                End If
                If Not Session.findById("wnd[1]/usr/txtMULTI_LOGON_TEXT", False) Is Nothing Then
                    If Not Session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2", False) Is Nothing Then
                        Session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").Select()
                    Else
                        Session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").Select()
                    End If
                End If
                If Not Session.findById("wnd[1]", False) Is Nothing Then Session.findById("wnd[1]").sendVKey(0)
                If Session.ActiveWindow.Text = "SAP" Then
                    Session.findById("wnd[1]/tbar[0]/btn[12]").Press()
                    Exit Sub
                End If
            Loop
            If Session.ActiveWindow.Text Like "SAP Easy Access*" Then
                LI = True
            End If
        Catch ex As Exception
            Sts = ex.Message
        End Try

    End Sub

    Public Sub Close()

        If Not FindById("wnd[1]/tbar[0]/btn[12]") Is Nothing Then
            Dim C = FindById("wnd[1]/tbar[0]/btn[12]")
            If Not C Is Nothing Then
                C.Press()
            End If
        Else
            If Not ActiveWindow() Is Nothing Then
                ActiveWindow.Close()
                If Not FindById("wnd[1]/usr/btnSPOP-OPTION1") Is Nothing Then
                    FindById("wnd[1]/usr/btnSPOP-OPTION1").press()
                End If
            End If
        End If

    End Sub

    Public Function FindById(ByVal ID As String) As Object

        Try
            FindById = Session.findById(ID, False)
        Catch ex As Exception
            FindById = Nothing
        End Try

    End Function

    Public Function FindByNameEx(ByVal Name As String, ByVal Type As Long) As Object

        Try
            FindByNameEx = Session.ActiveWindow.FindByNameEx(Name, Type)
        Catch ex As Exception
            FindByNameEx = Nothing
        End Try

    End Function

    Public Function FindAllByNameEx(ByVal Name As String, ByVal Type As Long) As Object

        Try
            FindAllByNameEx = Session.ActiveWindow.FindAllByNameEx(Name, Type)
        Catch ex As Exception
            FindAllByNameEx = Nothing
        End Try

    End Function

    Public Function FindByText(Search As String) As Object

        FindByText = Nothing
        Try
            For Each Children As Object In Session.findbyid("wnd[0]/usr/").children
                If Children.Text = Search Then
                    FindByText = Children
                    Exit For
                End If
            Next
        Catch ex As Exception
        End Try

    End Function

    Public Function ActiveWindow() As Object

        ActiveWindow = Session.ActiveWindow

    End Function

    Public Function GuiFocus() As Object

        GuiFocus = ActiveWindow.GuiFocus

    End Function

    Public Sub SendVKey(ByVal Code As Integer)

        Session.ActiveWindow.SendVKey(Code)

    End Sub

    Public Sub StartTransaction(ByVal Code As String)

        Session.StartTransaction(Code)

    End Sub

    Public Sub SendCommand(ByVal Code As String)

        Session.SendCommand(Code)

    End Sub

    Public Function DisplayPO(ByVal PO As String) As Boolean


        If Not LI Then
            DisplayPO = False
            Exit Function
        End If

        DisplayPO = True
        StartTransaction("me23n")
        FindByNameEx("btn[17]", 40).Press()
        FindByNameEx("MEPO_SELECT-EBELN", 32).Text = PO
        FindByNameEx("btn[0]", 40).Press()
        If StatusBarText <> "" Then
            DisplayPO = False
        End If

    End Function

    Public Function ChangePO(ByVal PO As String) As Boolean

        ChangePO = True
        If Not DisplayPO(PO) Then
            ChangePO = False
        Else
            FindByNameEx("btn[7]", 40).Press()
            If StatusBarText <> "" AndAlso StatusBarText <> "Text contains formatting -> SAPscript editor" Then
                ChangePO = False
            End If
        End If

    End Function

    Public Function GetConnString(ByVal Box As String) As String

        GetConnString = Nothing
        Servers.ReadXml(New IO.StringReader(My.Resources.Servers))
        Dim FR As DataRow() = Servers.Select("Box = '" & Box & "'")
        If FR.Count > 0 Then
            Dim R As String = "/R/*/G/" & FR(0)("LogonGroup") & "/M/" & FR(0)("MessageServer")
            GetConnString = R.Replace("*", Box)
        End If

    End Function

    Public Sub ArrayToClipboard(ByVal A() As String)

        Dim S As String = Nothing
        Dim C As String = Nothing
        My.Computer.Clipboard.Clear()
        For Each S In A
            C = C & S & Chr(13) & Chr(10)
        Next
        My.Computer.Clipboard.SetText(C)

    End Sub

    Public Sub TableToClipboard(ByVal DT As DataTable, ByVal ColumnIndex As Integer)

        Dim S As String = Nothing
        Dim DR As DataRow

        My.Computer.Clipboard.Clear()
        For Each DR In DT.Rows
            S = S & DR(ColumnIndex) & Chr(13) & Chr(10)
        Next
        My.Computer.Clipboard.SetText(S)

    End Sub

    Public Sub DRArrayToClipboard(ByVal DRA() As DataRow, ByVal ColumnIndex As Object)

        Dim S As String = Nothing
        Dim DR As DataRow

        My.Computer.Clipboard.Clear()
        For Each DR In DRA
            S = S & DR(ColumnIndex) & Chr(13) & Chr(10)
        Next
        My.Computer.Clipboard.SetText(S)

    End Sub

    Private Sub Enable_GUI_Theme()

        Dim PN As String = System.Diagnostics.Process.GetCurrentProcess.ProcessName
        Dim RV = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\SAP\General\Applications\" & PN, "Enjoy", Nothing)
        If RV Is Nothing Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\SAP\General\Applications\" & PN, "Enjoy", "On")
        End If

    End Sub

    Private Function GetSSOConnString(ByVal Box As String) As String

        GetSSOConnString = ""
        Select Case Box
            Case "L7P"
                GetSSOConnString = "L7P LA TS Prod - SSO"
            Case "L6A"
                GetSSOConnString = "L6A LA SC Acc - SSO"
            Case "L7A"
                GetSSOConnString = "L7A TS Acceptance - SSO"
            Case "L6P"
                GetSSOConnString = "L6P LA SC  Prod - SSO"
            Case "N6P"
                GetSSOConnString = "N6P NA Prod- SSO"
            Case "N6A"
                GetSSOConnString = "N6A NA SC Acc - SSO"
            Case "F6P"
                GetSSOConnString = "F6P EU SC Prod - SSO"
            Case "ANP"
                GetSSOConnString = "ANP NEA Prod(JP) - SSO"
            Case "A6P"
                GetSSOConnString = "A6P SC Prod(EN) - SSO"
            Case "GBP"
                GetSSOConnString = "GBP GCM Production- SSO"
            Case "G4P"
                GetSSOConnString = "G4P GCF/Cons Prod- SSO"
        End Select

    End Function

End Class

Public Class SAP_Faxing_Report

    Public Data As DataTable = Nothing
    Private GUI As SAPGUI
    Private FN As String = "Faxing.txt"

    Sub New(ByVal Box As String, ByVal User As String, ByVal Password As String, Optional ByVal Area As String = Nothing)

        GUI = New SAPGUI(Box, User, Password)
        If GUI.LoggedIn Then
            GUI.StartTransaction("Y_KLD_31001497")
            GUI.FindByNameEx("S_TRANS-LOW", 32).Text = ""
            GUI.FindByNameEx("S_WERKS-LOW", 32).Text = "*"
            GUI.FindByNameEx("%_S_EBELN_%_APP_%-VALU_PUSH", 40).press()
            GUI.FindByNameEx("btn[24]", 40).press()
            GUI.FindByNameEx("btn[8]", 40).press()
            GUI.FindByNameEx("btn[8]", 40).press()
            GUI.SendCommand("%pc")
            GUI.FindAllByNameEx("SPOPLI-SELFLAG", 41).Item(1).Select()
            GUI.FindByNameEx("btn[0]", 40).press()
            GUI.FindByNameEx("DY_PATH", 32).Text = My.Computer.FileSystem.SpecialDirectories.Temp
            GUI.FindByNameEx("DY_FILENAME", 32).Text = FN
            GUI.FindByNameEx("btn[11]", 40).press()
        End If
        GUI.Close()

        Data = New DataTable
        Data.Columns.Add("PO", System.Type.GetType("System.String"))
        Data.Columns.Add("Recno", System.Type.GetType("System.String"))
        Data.Columns.Add("Transm_Date", System.Type.GetType("System.DateTime"))
        Data.Columns.Add("Transm_Time", System.Type.GetType("System.String"))
        Data.Columns.Add("Fax", System.Type.GetType("System.String"))
        Data.Columns.Add("Success", System.Type.GetType("System.Boolean"))
        Data.Columns.Add("Message", System.Type.GetType("System.String"))
        If Not Area Is Nothing Then
            Data.Columns.Add("Area", System.Type.GetType("System.String"))
        End If
        Data.Columns.Add("SAP", System.Type.GetType("System.String"))

        Dim FileReader As New System.IO.StreamReader(My.Computer.FileSystem.SpecialDirectories.Temp & "\" & FN)
        Dim S As String
        Dim ExitLoop As Boolean = False
        Dim W As Array
        Dim DR As DataRow

        Do
            S = FileReader.ReadLine
            W = Split(S, Chr(9))
            If W.Length > 0 AndAlso W(0) = " Message#" Then
                S = FileReader.ReadLine
                ExitLoop = True
            End If
        Loop Until ExitLoop

        Do
            S = FileReader.ReadLine
            If Not S Is Nothing Then
                W = Split(S, Chr(9))
                DR = Data.NewRow
                DR("PO") = W(14)
                DR("Recno") = CDbl(W(0)).ToString
                DR("Transm_Date") = CDate(W(9))
                DR("Transm_Time") = W(10)
                DR("Fax") = W(13)
                If W(2) = "Successful" Then
                    DR("Success") = True
                Else
                    DR("Success") = False
                End If
                DR("Message") = W(3)
                If Not Area Is Nothing Then
                    DR("Area") = Area
                End If
                DR("SAP") = Box
                Data.Rows.Add(DR)
            End If
        Loop Until S Is Nothing

    End Sub

End Class

Public Class Refaxer

    Private IDN() As String = Nothing
    Private Box As String
    Private User As String
    Private Password As String

    Private Structure BW_Args
        Public Box As String
        Public User As String
        Public Password As String
        Public IDN() As String
    End Structure

    Sub New(ByVal ABox As String, ByVal AUser As String, ByVal APassword As String)

        Box = ABox
        User = AUser
        Password = APassword

    End Sub

    Public Sub IncludePO(ByVal Number As String)

        If IDN Is Nothing Then
            ReDim IDN(0)
        Else
            ReDim Preserve IDN(UBound(IDN) + 1)
        End If

        IDN(UBound(IDN)) = Number

    End Sub

    Public Sub Execute()

        Dim A As New BW_Args
        A.Box = Box
        A.User = User
        A.Password = Password
        A.IDN = IDN
        FaxScript(A)

    End Sub

    Private Sub FaxScript(ByVal A As BW_Args)

        Dim PO As String
        Dim TT As Object
        Dim I As Integer
        Dim EM As String

        Dim Session As New SAPGUI(A.Box, A.User, A.Password)

        If Session.LoggedIn Then
            For Each PO In A.IDN
                If Session.ChangePO(PO) Then
                    Try
                        Session.FindById("wnd[0]/tbar[1]/btn[21]").Press()
                        TT = Session.FindAllByNameEx("DNAST-KSCHL", 32)
                        I = 0
                        Do While TT.ElementAt(I).Text <> "" And I <= TT.Count - 1
                            If Not Session.FindById("wnd[0]/usr/tblSAPDV70ATC_NAST3/lblDV70A-STATUSICON[0," & I & "]") Is Nothing AndAlso _
                            Session.FindById("wnd[0]/usr/tblSAPDV70ATC_NAST3/lblDV70A-STATUSICON[0," & I & "]").ToolTip = "Not processed" Then
                                Session.FindByNameEx("SAPDV70ATC_NAST3", 80).getAbsoluteRow(I).selected = True
                                Session.FindByNameEx("btn[18]", 40).press()
                                If Not Session.FindById("wnd[1]/tbar[0]/btn[0]") Is Nothing Then
                                    Session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
                                End If
                                TT = Session.FindAllByNameEx("DNAST-KSCHL", 32)
                            Else
                                I += 1
                            End If
                        Loop
                        If I < TT.Count - 1 Then
                            TT.ElementAt(I).Text = "NNXX"
                            Session.FindAllByNameEx("NAST-NACHA", 34).ElementAt(I).Key = "2"
                            Session.FindById("wnd[0]/tbar[1]/btn[2]").Press()
                            Session.FindById("wnd[1]/tbar[0]/btn[2]").Press()
                            If Session.FindById("wnd[1]/usr/txtNAST-TELFX").Text <> "" AndAlso Session.FindById("wnd[1]/usr/ctxtNAST-TLAND").Text = "US" Then
                                Session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
                                Session.FindById("wnd[0]/tbar[1]/btn[5]").Press()
                                Session.FindByNameEx("NAST-VSZTP", 34).Key = "4"
                                Session.FindById("wnd[0]/tbar[0]/btn[3]").Press()
                                Session.FindById("wnd[0]/tbar[0]/btn[11]").Press()  'SAVE BUTTON
                                If Not Session.FindById("wnd[1]/usr/btnSPOP-VAROPTION1") Is Nothing Then
                                    Session.FindById("wnd[1]/usr/btnSPOP-VAROPTION1").Press()
                                End If
                            Else
                                Session.FindById("wnd[1]/tbar[0]/btn[12]").Press()
                            End If

                        End If
                    Catch ex As Exception
                        EM = ex.Message
                    End Try
                End If
            Next
        End If
        Session.Close()

    End Sub

End Class

Public Class BI_Release

    Private BIDT As DataTable
    Private Session As SAPGUI = Nothing
    Private Box As String
    Private User As String
    Private Password As String
    Private FY As String

    Public Event Notify(ByVal Box As String, ByVal IR As String, ByVal Result As String)
    Public Event MSA_Notify(ByVal Box As String, ByVal IR As String, ByVal ILI As String, ByVal PO As String, ByVal POLI As String, ByVal Result As String)

    Sub New(ByVal ABox As String, ByVal AUser As String, ByVal APassword As String, ByVal FiscalYear As String)

        Box = ABox
        User = AUser
        Password = APassword
        FY = FiscalYear

    End Sub

    Public Sub LoadDataTable(ByVal DT As DataTable)

        BIDT = New DataTable
        Dim Reader As New DataTableReader(DT)
        BIDT.Load(Reader)

    End Sub

    Public Sub Execute()

        If BIDT.Rows.Count <= 0 Then Exit Sub

        Dim BIDR() As DataRow

        Session = New SAPGUI(Box, User, Password)

        BIDR = BIDT.Select("Manual = ''")
        If BIDR.Length > 0 Then
            If BIDT.Columns.Contains("InvoiceLineItem") Then
                Release_BI_Item_PQ(BIDR)
            Else
                Release_BI_Price(BIDR)
            End If
        End If

        BIDR = BIDT.Select("Manual = 'X'")
        If BIDR.Length > 0 Then
            Release_D(BIDR)
        End If

        Session.Close()

    End Sub

    Private Sub Run_ZMR0(ByVal BIDR() As DataRow, ByVal PRC_Flag As Boolean, ByVal QTY_Flag As Boolean, ByVal MNL_Flag As Boolean)

        Session.StartTransaction("ZMR0")
        Session.FindByNameEx("S_BUKRS-LOW", 32).Text = "*"
        Session.FindByNameEx("S_GJAHR-LOW", 31).Text = FY
        Session.FindByNameEx("S_GJAHR-LOW", 31).SetFocus()
        Session.SendVKey(2)
        Session.FindByNameEx("shell", 122).SelectedRows = "2"
        Session.FindByNameEx("btn[0]", 40).Press()
        Session.FindByNameEx("P_COUNT", 42).Selected = False
        Session.FindByNameEx("P_INPT", 42).Selected = False
        Session.FindByNameEx("P_IMAGE", 42).Selected = False
        Session.FindByNameEx("P_SPGRP", 42).Selected = PRC_Flag
        Session.FindByNameEx("P_SPGRM", 42).Selected = QTY_Flag
        Session.FindByNameEx("P_MANL", 42).Selected = MNL_Flag
        Session.DRArrayToClipboard(BIDR, "Vendor")
        Session.FindByNameEx("%_S_LIFNR_%_APP_%-VALU_PUSH", 40).Press()
        Session.FindByNameEx("btn[24]", 40).Press()
        Session.FindByNameEx("btn[8]", 40).Press()
        Session.FindByNameEx("btn[8]", 40).Press()

    End Sub

    Private Sub Release_BI_Item_PQ(ByVal BIDR() As DataRow)

        Dim DR As DataRow
        If Session.LoggedIn Then
            Try
                Run_ZMR0(BIDR, True, True, False)
                Dim Proc As Boolean
                Dim RS As String = Nothing
                For Each DR In BIDR
                    Proc = False
                    If Find_IR_Item_Price(DR("InvoiceNumber"), DR("InvoiceLineItem")) Then
                        RS = Release_Selected(DR("SAP_Release_Code"), DR("SAP_Comments"))
                        Proc = True
                    End If
                    If Find_IR_Item_Quantity(DR("InvoiceNumber"), DR("InvoiceLineItem")) Then
                        RS = Release_Selected(DR("SAP_Release_Code"), DR("SAP_Comments"))
                        Proc = True
                    End If
                    If Proc Then
                        RaiseEvent MSA_Notify(Box, DR("InvoiceNumber"), DR("InvoiceLineItem"), DR("PurchaseDoc"), DR("POLineItem"), RS)
                    Else
                        RaiseEvent MSA_Notify(Box, DR("InvoiceNumber"), DR("InvoiceLineItem"), DR("PurchaseDoc"), DR("POLineItem"), "IR Already Released!")
                    End If
                Next
            Catch ex As Exception
                Dim S As String
                S = ex.Message
            End Try
        Else
            For Each DR In BIDR
                RaiseEvent MSA_Notify(Box, DR("InvoiceNumber"), DR("InvoiceLineItem"), DR("PurchaseDoc"), DR("POLineItem"), "SAP Login Failed!")
            Next
        End If

    End Sub

    Private Sub Release_BI_Price(ByVal BIDR() As DataRow)

        Dim DR As DataRow
        If Session.LoggedIn Then
            Try
                Run_ZMR0(BIDR, True, False, False)
                Dim Proc As Boolean
                Dim RS As String = Nothing
                For Each DR In BIDR
                    Proc = False
                    Do While FindIR_Price(DR("InvoiceNumber"))
                        If Not DBNull.Value.Equals(DR("SAP_Release_Code")) Then
                            RS = Release_Selected(DR("SAP_Release_Code"), DR("SAP_Comments"))
                        Else
                            RS = Release_Selected("4", "Price discrepancy under tolerance")
                        End If
                        Proc = True
                    Loop
                    If Proc Then
                        RaiseEvent Notify(Box, DR("InvoiceNumber"), RS)
                    Else
                        RaiseEvent Notify(Box, DR("InvoiceNumber"), "IR Already Released!")
                    End If
                Next
            Catch ex As Exception
                Dim S As String
                S = ex.Message
            End Try
        Else
            For Each DR In BIDR
                RaiseEvent Notify(Box, DR("InvoiceNumber"), "SAP Login Failed!")
            Next
        End If

    End Sub

    Private Sub Release_D(ByVal BIDR() As DataRow)

        Dim DR As DataRow
        Dim RC As String

        If Session.LoggedIn Then
            Run_ZMR0(BIDR, False, False, True)
            Dim Proc As Boolean
            Dim RS As String = Nothing
            For Each DR In BIDR
                Proc = False
                If FindIR_D(DR("InvoiceNumber")) Then
                    If Box = "N6P" Then
                        RC = "26"
                    Else
                        RC = "44"
                    End If
                    RS = Release_Selected(RC, "Manual Block. No PO\Invoice discrepancy")
                    Proc = True
                End If
                If Proc Then
                    RaiseEvent Notify(Box, DR("InvoiceNumber"), RS)
                Else
                    RaiseEvent Notify(Box, DR("InvoiceNumber"), "IR Already Released!")
                End If
            Next
        Else
            For Each DR In BIDR
                RaiseEvent Notify(Box, DR("InvoiceNumber"), "SAP Login Failed!")
            Next
        End If

    End Sub

    Private Function FindIR_Price(ByVal IR As String) As Boolean

        FindIR_Price = False

        If Session.FindByNameEx("btn[13]", 40) Is Nothing Then Exit Function

        Session.SendVKey(71)
        Session.FindByNameEx("RSYSF-STRING", 31).Text = IR
        Session.FindByNameEx("SCAN_STRING-START", 42).Selected = False
        Session.FindByNameEx("btn[0]", 40).Press()

        If Not Session.ActiveWindow.Text = "Find" Then
            Session.FindByNameEx("btn[0]", 40).Press()
            Session.FindByNameEx("btn[12]", 40).Press()
            Exit Function
        End If

        Dim I As Integer = 2
        Dim Found As Boolean = False
        Do While Not Session.FindById("wnd[2]/usr/lbl[48," & I & "]") Is Nothing And Not Found
            If Session.FindById("wnd[2]/usr/lbl[48," & I & "]").Text = "X" Or Session.FindById("wnd[2]/usr/lbl[48," & I & "]").Text = "R" Then
                Found = True
                Session.FindById("wnd[2]/usr/lbl[48," & I & "]").SetFocus()
            Else
                I += 1
            End If
        Loop

        If Found Then
            FindIR_Price = True
            Session.FindByNameEx("btn[2]", 40).Press()
        Else
            Session.FindByNameEx("btn[12]", 40).Press()
            Session.FindByNameEx("btn[12]", 40).Press()
        End If

    End Function

    Private Function FindIR_D(ByVal IR As String) As Boolean

        FindIR_D = False

        If Session.FindByNameEx("btn[13]", 40) Is Nothing Then Exit Function

        Session.SendVKey(71)
        Session.FindByNameEx("RSYSF-STRING", 31).Text = IR
        Session.FindByNameEx("SCAN_STRING-START", 42).Selected = False
        Session.FindByNameEx("btn[0]", 40).Press()

        If Not Session.ActiveWindow.Text = "Find" Then
            Session.FindByNameEx("btn[0]", 40).Press()
            Session.FindByNameEx("btn[12]", 40).Press()
            Exit Function
        End If

        Session.FindById("wnd[2]/usr/lbl[5,2]").SetFocus()
        Session.SendVKey(2)
        Dim I As String = Left(Right(Session.GuiFocus.ID, 3), 2)
        If Left(I, 1) = "," Then I = Right(I, 1)
        I = (Val(I) + 1).ToString
        If Session.FindById("wnd[0]/usr/lbl[117," & I & "]").Text = "X" Or Session.FindById("wnd[0]/usr/lbl[117," & I & "]").Text = "R" Then
            Session.FindById("wnd[0]/usr/lbl[117," & I & "]").Setfocus()
            FindIR_D = True
        End If

    End Function

    Private Function Release_Selected(ByVal RCode As String, ByVal RMessage As String) As String

        Release_Selected = Nothing
        Try
            Session.SendVKey(2)
            Session.FindByNameEx("ZMINVMR-MR_CODE", 32).text = RCode
            Session.FindByNameEx("ZMINVMR-MR_CUST_TEXT", 31).text = RMessage
            Session.FindByNameEx("btn[11]", 40).Press()
            Session.FindByNameEx("btn[3]", 40).Press()
            Session.FindByNameEx("btn[13]", 40).Press()
            Session.FindByNameEx("btn[0]", 40).Press()
            Release_Selected = Session.StatusBarText
        Catch ex As Exception
            Dim S As String
            S = ex.Message
        End Try

    End Function

    Private Function Find_IR_Item_Price(ByVal IR As String, ByVal Item As String) As Boolean

        Find_IR_Item_Price = False

        If Session.FindByNameEx("btn[13]", 40) Is Nothing Then Exit Function

        Session.SendVKey(71)
        Session.FindByNameEx("RSYSF-STRING", 31).Text = IR
        Session.FindByNameEx("SCAN_STRING-START", 42).Selected = False
        Session.FindByNameEx("btn[0]", 40).Press()

        If Not Session.ActiveWindow.Text = "Find" Then
            Session.FindByNameEx("btn[0]", 40).Press()
            Session.FindByNameEx("btn[12]", 40).Press()
            Exit Function
        End If

        Dim I As Integer = 2
        Dim Found As Boolean = False
        Dim ItemFound As Boolean = False
        Do While Not Session.FindById("wnd[2]/usr/lbl[48," & I & "]") Is Nothing And Not ItemFound
            If Val(Session.FindById("wnd[2]/usr/lbl[16," & I & "]").Text) = Val(Item) Then
                ItemFound = True
                If Session.FindById("wnd[2]/usr/lbl[48," & I & "]").Text = "X" Or Session.FindById("wnd[2]/usr/lbl[48," & I & "]").Text = "R" Then
                    Found = True
                    Session.FindById("wnd[2]/usr/lbl[48," & I & "]").SetFocus()
                Else
                    I += 1
                End If
            End If
        Loop

        If Found Then
            Find_IR_Item_Price = True
            Session.FindByNameEx("btn[2]", 40).Press()
        Else
            Session.FindByNameEx("btn[12]", 40).Press()
            Session.FindByNameEx("btn[12]", 40).Press()
        End If

    End Function

    Private Function Find_IR_Item_Quantity(ByVal IR As String, ByVal Item As String) As Boolean

        Find_IR_Item_Quantity = False

        If Session.FindByNameEx("btn[13]", 40) Is Nothing Then Exit Function

        Session.SendVKey(71)
        Session.FindByNameEx("RSYSF-STRING", 31).Text = IR
        Session.FindByNameEx("SCAN_STRING-START", 42).Selected = False
        Session.FindByNameEx("btn[0]", 40).Press()

        If Not Session.ActiveWindow.Text = "Find" Then
            Session.FindByNameEx("btn[0]", 40).Press()
            Session.FindByNameEx("btn[12]", 40).Press()
            Exit Function
        End If

        Dim I As Integer = 2
        Dim Found As Boolean = False
        Dim ItemFound As Boolean = False
        Do While Not Session.FindById("wnd[2]/usr/lbl[46," & I & "]") Is Nothing And Not ItemFound
            If Val(Session.FindById("wnd[2]/usr/lbl[16," & I & "]").Text) = Val(Item) Then
                ItemFound = True
                If Session.FindById("wnd[2]/usr/lbl[46," & I & "]").Text = "X" Or Session.FindById("wnd[2]/usr/lbl[46," & I & "]").Text = "R" Then
                    Found = True
                    Session.FindById("wnd[2]/usr/lbl[46," & I & "]").SetFocus()
                Else
                    I += 1
                End If
            End If
        Loop

        If Found Then
            Find_IR_Item_Quantity = True
            Session.FindByNameEx("btn[2]", 40).Press()
        Else
            Session.FindByNameEx("btn[12]", 40).Press()
            Session.FindByNameEx("btn[12]", 40).Press()
        End If

    End Function

End Class

Public Class DRT_SL_Maintain

    Private Session As SAPGUI = Nothing
    Private Const SAP_BOX As String = "L6P"

    Public ReadOnly Property IsReady() As Boolean

        Get
            IsReady = False
            If Not Session Is Nothing Then IsReady = Session.LoggedIn
        End Get

    End Property

    Public Sub New(ByVal User As String, ByVal Password As String)

        Session = New SAPGUI(SAP_BOX, User, Password)

    End Sub

    Public Sub Close()

        If Not Session Is Nothing Then
            Session.Close()
        End If

    End Sub

    Public Function Execute(ByVal OA As String, ByVal ItemList As DataTable) As DataTable

        Dim R As New DataTable
        Dim RR As DataRow

        R.Columns.Add("OA", Type.GetType("System.String"))
        R.Columns.Add("Items", Type.GetType("System.String"))
        R.Columns.Add("Item", Type.GetType("System.String"))
        R.Columns.Add("Proc_Time", Type.GetType("System.DateTime"))
        R.Columns.Add("Result", Type.GetType("System.String"))
        R.PrimaryKey = New DataColumn() {R.Columns(0), R.Columns(1), R.Columns(2)}
        R.TableName = "FP_SL_Maintain"

        If Not Session.LoggedIn Then
            For Each GCas As DataRow In ItemList.Rows
                RR = R.NewRow
                RR("OA") = OA
                RR("Items") = GCas(0)
                RR("Item") = ""
                RR("Proc_Time") = My.Computer.Clock.LocalTime
                RR("Result") = "SAP Login Failed"
                Try
                    R.Rows.Add(RR)
                Catch ex As Exception
                End Try
            Next
        Else
            Dim VStart = "07/01/" & My.Computer.Clock.LocalTime.Year
            Dim VEnd = "12/31/" & My.Computer.Clock.LocalTime.AddYears(1).Year
            Session.StartTransaction("ME32K")
            Session.FindByNameEx("RM06E-EVRTN", 32).Text = OA
            Session.SendVKey(0)
            If Session.StatusBarMessageType = "E" Then
                For Each GCas As DataRow In ItemList.Rows
                    RR = R.NewRow
                    RR("OA") = OA
                    RR("Items") = GCas(0)
                    RR("Item") = ""
                    RR("Proc_Time") = My.Computer.Clock.LocalTime
                    RR("Result") = Session.StatusBarText
                    Try
                        R.Rows.Add(RR)
                    Catch ex As Exception
                    End Try
                Next
            Else
                For Each GCas As DataRow In ItemList.Rows
                    For Each Item As String In GCas(0).Split("/")
                        RR = R.NewRow
                        RR("OA") = OA
                        RR("Items") = GCas(0)
                        RR("Item") = Item.Trim
                        RR("Proc_Time") = My.Computer.Clock.LocalTime
                        Session.FindByNameEx("RM06E-EBELP", 31).Text = Item.Trim
                        Session.SendVKey(0)
                        If Session.FindByNameEx("SAPMM06ETC_0220", 80).GetCell(0, 0).Text = Item.Trim Then
                            Session.FindByNameEx("SAPMM06ETC_0220", 80).Rows(0).Selected = True
                            Session.FindById("wnd[0]/mbar/menu[3]/menu[4]").select()
                            If Not Session.FindById("wnd[1]/tbar[0]/btn[0]") Is Nothing Then
                                Session.FindById("wnd[1]/tbar[0]/btn[0]").press()
                            End If
                            If Session.FindById("wnd[1]/tbar[0]/btn[0]") Is Nothing Then
                                If Session.FindByNameEx("EORD-VDATU", 32).Text = "" Then
                                    Session.FindByNameEx("EORD-VDATU", 32).Text = VStart
                                End If
                                Session.FindByNameEx("EORD-BDATU", 32).Text = VEnd
                                If Session.FindByNameEx("EORD-AUTET", 32).Text <> "1" Then
                                    Session.FindByNameEx("EORD-AUTET", 32).Text = "1"
                                End If
                                Session.FindByNameEx("btn[3]", 40).Press()
                                Do While Session.StatusBarText <> ""
                                    If Session.StatusBarMessageType <> "E" Then
                                        Session.SendVKey(0)
                                    Else
                                        RR("Result") = Session.StatusBarText
                                        Session.FindByNameEx("btn[15]", 40).Press()
                                    End If
                                Loop
                                If DBNull.Value.Equals(RR("Result")) Then RR("Result") = "Source List Updated"
                            Else
                                Session.FindById("wnd[1]/tbar[0]/btn[0]").press()
                                RR("Result") = "Item Blocked/Deleted"
                            End If
                        Else
                            RR("Result") = "Item Not Found"
                        End If
                        Try
                            R.Rows.Add(RR)
                        Catch ex As Exception
                        End Try
                    Next
                Next
                Session.FindByNameEx("btn[11]", 40).Press()
                If Not Session.FindById("wnd[1]/usr/btnSPOP-OPTION1") Is Nothing Then
                    Session.FindById("wnd[1]/usr/btnSPOP-OPTION1").press()
                End If
                If Not Session.FindById("wnd[1]/usr/btnBUTTON_1") Is Nothing Then
                    Session.FindById("wnd[1]/usr/btnBUTTON_1").press()
                End If
            End If
        End If

        Execute = R

    End Function

End Class

Public Class ADM_SL_Maintain

    Private Session As SAPGUI = Nothing
    Private Const SAP_BOX As String = "L6P"
    Private OA As String

    Public ReadOnly Property IsReady() As Boolean

        Get
            IsReady = False
            If Not Session Is Nothing Then IsReady = Session.LoggedIn
        End Get

    End Property

    Public Sub New()

        Session = New SAPGUI(SAP_BOX)

    End Sub

    Public Sub Close()

        If Not Session Is Nothing Then
            Session.Close()
        End If

    End Sub

    Private Function OA_Match(ByVal DR As DataRow) As Boolean

        If DR("Contract") = OA Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function SL_Maintain(ByVal MDT() As DataRow) As DataTable

        Dim R As New DataTable("ADM_SL_Maintenance")
        Dim RR As DataRow

        R.Columns.Add("Request", Type.GetType("System.Int32"))
        R.Columns.Add("OA", Type.GetType("System.String"))
        R.Columns.Add("Item", Type.GetType("System.String"))
        R.Columns.Add("Proc_Time", Type.GetType("System.DateTime"))
        R.Columns.Add("Result", Type.GetType("System.String"))
        R.PrimaryKey = New DataColumn() {R.Columns(0), R.Columns(1), R.Columns(2)}

        If Not Session.LoggedIn Then
            For Each DR As DataRow In MDT
                RR = R.NewRow
                RR("Request") = DR("Request")
                RR("OA") = DR("Contract")
                RR("Item") = DR("Item")
                RR("Proc_Time") = My.Computer.Clock.LocalTime
                RR("Result") = "SAP Login Failed"
                Try
                    R.Rows.Add(RR)
                Catch ex As Exception
                End Try
            Next
        Else
            Dim QOA = From RMDT In MDT Group By OA = RMDT("Contract") Into G = Group Select OA
            For Each GOA As String In QOA
                OA = GOA
                Session.StartTransaction("ME32")
                Session.FindByNameEx("RM06E-EVRTN", 32).Text = OA
                Session.SendVKey(0)
                If Session.StatusBarMessageType = "E" Then
                    For Each MDR As DataRow In Array.FindAll(MDT, AddressOf OA_Match)
                        RR = R.NewRow
                        RR("Request") = MDR("Request")
                        RR("OA") = OA
                        RR("Item") = MDR("Item")
                        RR("Proc_Time") = My.Computer.Clock.LocalTime
                        RR("Result") = Session.StatusBarText
                        Try
                            R.Rows.Add(RR)
                        Catch ex As Exception
                        End Try
                    Next
                Else
                    For Each MDR As DataRow In Array.FindAll(MDT, AddressOf OA_Match)
                        RR = R.NewRow
                        RR("Request") = MDR("Request")
                        RR("OA") = OA
                        RR("Item") = MDR("Item")
                        RR("Proc_Time") = My.Computer.Clock.LocalTime
                        For Each Item As String In MDR("Item").Split("/")
                            Session.FindByNameEx("RM06E-EBELP", 31).Text = Item.Trim
                            Session.SendVKey(0)
                            If Session.FindByNameEx("SAPMM06ETC_0220", 80).GetCell(0, 0).Text = Item.Trim Then
                                Session.FindByNameEx("SAPMM06ETC_0220", 80).Rows(0).Selected = True
                                If OA.StartsWith("5") Then
                                    Session.FindById("wnd[0]/mbar/menu[3]/menu[6]").select()
                                Else
                                    Session.FindById("wnd[0]/mbar/menu[3]/menu[4]").select()
                                End If

                                If IsDate(Session.FindByNameEx("EORD-BDATU", 32).Text) AndAlso IsDate(MDR("Validity_end")) Then
                                    If CDate(Session.FindByNameEx("EORD-BDATU", 32).Text).CompareTo(CDate(MDR("Validity_end"))) < 0 Then
                                        Session.FindByNameEx("EORD-BDATU", 32).Text = MDR("Validity_end")
                                        If Not DBNull.Value.Equals(RR("Result")) Then RR("Result") = RR("Result") & " \ "
                                        RR("Result") = RR("Result") & "Item " & Item & " : Source List Updated"
                                    Else
                                        If Not DBNull.Value.Equals(RR("Result")) Then RR("Result") = RR("Result") & " \ "
                                        RR("Result") = RR("Result") & "Item " & Item & " : No Need to Update. Validity ends on " & Session.FindByNameEx("EORD-BDATU", 32).Text
                                    End If
                                Else
                                    If Not DBNull.Value.Equals(RR("Result")) Then RR("Result") = RR("Result") & " \ "
                                    RR("Result") = RR("Result") & "Item " & Item & " : Dates could not be compared (SL: '" & Session.FindByNameEx("EORD-BDATU", 32).Text & "' Form: '" & MDR("Validity_end") & "')"
                                End If

                                Session.FindByNameEx("btn[3]", 40).Press()
                                Do While Session.StatusBarText <> ""
                                    If Session.StatusBarMessageType <> "E" Then
                                        Do While Session.StatusBarText <> ""
                                            Session.SendVKey(0)
                                        Loop
                                    Else
                                        If Not DBNull.Value.Equals(RR("Result")) Then RR("Result") = RR("Result") & " \ "
                                        RR("Result") = RR("Result") & "Item " & Item & " : " & Session.StatusBarText
                                        Session.FindByNameEx("btn[15]", 40).Press()
                                    End If
                                Loop

                            Else
                                If Not DBNull.Value.Equals(RR("Result")) Then RR("Result") = RR("Result") & " \ "
                                RR("Result") = RR("Result") & "Item " & Item & " Not Found"
                            End If
                            Try
                                R.Rows.Add(RR)
                            Catch ex As Exception
                            End Try
                        Next
                    Next
                    Session.FindByNameEx("btn[11]", 40).Press()
                    If Not Session.FindById("wnd[1]/usr/btnSPOP-OPTION1") Is Nothing Then
                        Session.FindById("wnd[1]/usr/btnSPOP-OPTION1").press()
                    End If
                    If Not Session.FindById("wnd[1]/usr/btnBUTTON_1") Is Nothing Then
                        Session.FindById("wnd[1]/usr/btnBUTTON_1").press()
                    End If
                End If
            Next
        End If

        SL_Maintain = R

    End Function

End Class

Public Class DRT_Inforecords

    Private Session As SAPGUI = Nothing
    Private Const SAP_BOX As String = "L6P"
    Private RS As String = Nothing
    Private VStart = "07/01/" & My.Computer.Clock.LocalTime.Year

    Public ReadOnly Property IsReady() As Boolean

        Get
            IsReady = False
            If Not Session Is Nothing Then IsReady = Session.LoggedIn
        End Get

    End Property

    Public ReadOnly Property Result() As String

        Get
            Result = RS
        End Get

    End Property

    Public ReadOnly Property IR_Number() As String

        Get
            IR_Number = Nothing
            If Not RS Is Nothing Then
                Dim FFP As Integer = RS.IndexOf("5")
                If FFP > 0 Then
                    Dim TN As String = RS.Substring(FFP, 10)
                    If IsNumeric(TN) Then
                        IR_Number = TN
                    End If
                End If
            End If
        End Get

    End Property

    Public Sub New(ByVal User As String, ByVal Password As String)

        Session = New SAPGUI(SAP_BOX, User, Password)

    End Sub

    Public Sub Close()

        If Not Session Is Nothing Then
            Session.Close()
        End If

    End Sub

    Public Sub Execute(ByVal Vendor As String, ByVal Material As String, ByVal Plant As String, ByVal Amount As String, ByVal ZHC2 As String, ByVal UOM As String)

        If Not IsReady Then
            RS = "Login Failed"
        Else
            If Not Create(Vendor, Material, Plant, Amount, ZHC2, UOM) Then
                If Not IR_Number Is Nothing Then
                    Update(IR_Number, Vendor, Material, Plant, Amount, ZHC2, UOM)
                End If
            End If
        End If

    End Sub

    Private Function Create(ByVal Vendor As String, ByVal Material As String, ByVal Plant As String, ByVal Amount As String, ByVal ZHC2 As String, ByVal UOM As String) As Boolean

        Create = False
        Session.StartTransaction("ME11")
        Session.FindByNameEx("EINA-LIFNR", 32).Text = Vendor
        Session.FindByNameEx("EINA-MATNR", 32).Text = Material
        Session.FindByNameEx("EINE-EKORG", 32).Text = "1159"
        Session.FindByNameEx("EINE-WERKS", 32).Text = Plant
        Session.SendVKey(0)
        Do While Session.StatusBarText <> ""
            If Session.StatusBarMessageType = "E" Then
                RS = Session.StatusBarText
                Exit Function
            Else
                Session.SendVKey(0)
            End If
        Loop
        Session.FindByNameEx("EINA-MAHN1", 31).Text = "1"
        Session.SendVKey(0)
        Do While Session.StatusBarText <> ""
            If Session.StatusBarMessageType = "E" Then
                RS = Session.StatusBarText
                Exit Function
            Else
                Session.SendVKey(0)
            End If
        Loop
        Session.FindByNameEx("EINE-APLFZ", 31).Text = "1"       'PDT
        Session.FindByNameEx("EINE-EKGRP", 32).Text = "278"     'PGrp
        Session.FindByNameEx("EINE-NORBM", 31).Text = "1"       'Std Qty
        Session.FindByNameEx("EINE-MINBM", 31).Text = "1"       'Min Qty
        Session.FindByNameEx("EINE-NETPR", 31).Text = "1"       'Net Price
        Session.SendVKey(0)
        Do While Session.StatusBarText <> ""
            If Session.StatusBarMessageType = "E" Then
                RS = Session.StatusBarText
                Exit Function
            Else
                Session.SendVKey(0)
            End If
        Loop
        Session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
        Session.FindByNameEx("RV13A-DATAB", 32).Text = VStart   'Valid From
        Dim GTC = Session.FindByNameEx("SAPMV13ATCTRL_D0201", 80)
        GTC.GetCell(0, 2).Text = Amount
        GTC.GetCell(0, 3).Text = "USD"
        GTC.GetCell(0, 4).Text = "1000"
        GTC.GetCell(0, 5).Text = UOM
        GTC.GetCell(1, 0).Text = "ZHC2"
        GTC.GetCell(1, 2).Text = ZHC2
        GTC.GetCell(1, 3).Text = "USD"
        GTC.GetCell(1, 4).Text = "1000"
        GTC.GetCell(1, 5).Text = UOM
        Session.FindById("wnd[0]/tbar[0]/btn[11]").Press()   'Save
        If Not Session.FindById("wnd[1]/tbar[0]/btn[0]") Is Nothing Then
            RS = Session.FindById("wnd[1]/usr/txtMESSTXT1").text
            Session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
            Create = True
        Else
            RS = Session.StatusBarText
            If Session.StatusBarMessageType <> "E" Then
                Create = True
            End If
        End If

    End Function

    Private Function Update(ByVal InfoRecord As String, ByVal Vendor As String, ByVal Material As String, ByVal Plant As String, ByVal Amount As String, ByVal ZHC2 As String, ByVal UOM As String) As Boolean

        Update = False
        Session.StartTransaction("ME12")
        Session.FindByNameEx("EINA-INFNR", 32).Text = InfoRecord
        Session.FindByNameEx("EINA-LIFNR", 32).Text = Vendor
        Session.FindByNameEx("EINA-MATNR", 32).Text = Material
        Session.FindByNameEx("EINE-EKORG", 32).Text = "1159"
        Session.FindByNameEx("EINE-WERKS", 32).Text = Plant
        Session.SendVKey(0)
        Do While Session.StatusBarText <> ""
            If Session.StatusBarMessageType = "E" Then
                RS = Session.StatusBarText
                Exit Function
            Else
                Session.SendVKey(0)
            End If
        Loop
        Session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
        Session.FindById("wnd[1]/tbar[0]/btn[7]").Press()
        Session.FindByNameEx("RV13A-DATAB", 32).Text = VStart   'Valid From
        Dim GTC = Session.FindByNameEx("SAPMV13ATCTRL_D0201", 80)
        GTC.GetCell(0, 2).Text = Amount
        GTC.GetCell(0, 3).Text = "USD"
        GTC.GetCell(0, 4).Text = "1000"
        GTC.GetCell(0, 5).Text = UOM
        GTC.GetCell(1, 0).Text = "ZHC2"
        GTC.GetCell(1, 2).Text = ZHC2
        GTC.GetCell(1, 3).Text = "USD"
        GTC.GetCell(1, 4).Text = "1000"
        GTC.GetCell(1, 5).Text = UOM
        Session.FindById("wnd[0]/tbar[0]/btn[11]").Press()   'Save

        If Not Session.FindById("wnd[1]/tbar[0]/btn[0]") Is Nothing Then
            RS = Session.FindById("wnd[1]/usr/txtMESSTXT1").text
            Session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
        Else
            If Not Session.FindById("wnd[1]/tbar[0]/btn[5]") Is Nothing Then
                Session.FindById("wnd[1]/tbar[0]/btn[5]").Press()
            End If
            RS = Session.StatusBarText
            If Session.StatusBarMessageType <> "E" Then
                Update = True
            End If
        End If

    End Function

End Class

Public Class Contract_DeleteItems

    Private Const SAP_BOX As String = "L6P"
    Private Session As SAPGUI = Nothing

    Public Sub New(ByVal User As String, ByVal Password As String)

        Session = New SAPGUI(SAP_BOX, User, Password)

    End Sub

    Public Function Execute(ByVal OA As String, ByVal DT As DataTable) As String

        Session.StartTransaction("ME32K")
        Session.FindByNameEx("RM06E-EVRTN", 32).Text = OA
        Session.SendVKey(0)
        For Each DR As DataRow In DT.Rows
            Session.FindByNameEx("RM06E-EBELP", 31).Text = DR("Item").ToString.Trim
            Session.SendVKey(0)
            If Session.FindByNameEx("SAPMM06ETC_0220", 80).GetCell(0, 0).Text = DR("Item").ToString.Trim Then
                Session.FindByNameEx("SAPMM06ETC_0220", 80).Rows(0).Selected = True
                Session.FindById("wnd[0]/tbar[1]/btn[14]").Press()
                If Not Session.FindById("wnd[1]/usr/btnSPOP-OPTION1") Is Nothing Then
                    Session.FindById("wnd[1]/usr/btnSPOP-OPTION1").Press()
                End If
            End If
            Do While Session.StatusBarText <> ""
                Session.SendVKey(0)
            Loop
        Next
        Session.FindById("wnd[0]/tbar[0]/btn[11]").Press()
        If Not Session.FindById("wnd[1]/usr/btnSPOP-OPTION1") Is Nothing Then
            Session.FindById("wnd[1]/usr/btnSPOP-OPTION1").Press()
        End If

        Execute = Session.StatusBarText

    End Function

End Class

Public Class PO_Condition_Creator

    Private Session As SAPGUI = Nothing
    Private RDT As DataTable = Nothing

    Public ReadOnly Property Result() As DataTable

        Get
            Result = RDT
        End Get

    End Property

    Public Sub New(ByVal Box As String)

        Session = New SAPGUI(Box)

    End Sub

    Public Sub Close()

        If Not Session Is Nothing Then
            Session.Close()
        End If

    End Sub

    Public Sub Execute(ByVal DT As DataTable)

        If Not Session.LoggedIn Then
            Exit Sub
        End If

        Dim TableControl
        Dim TableRow
        RDT = New DataTable

        RDT.Columns.Add("PO", Type.GetType("System.String"))
        RDT.Columns.Add("Result", Type.GetType("System.String"))

        For Each DR As DataRow In DT.Rows

            Session.ChangePO(DR("PURCH#DOC#"))
            For Item As Integer = 1 To Session.FindByNameEx("DYN_6000-LIST", 34).Entries.Count - 1
                Session.FindByNameEx("DYN_6000-LIST", 34).key = Item.ToString.PadLeft(4, " ")
                TableControl = Session.FindByNameEx("SAPLV69ATCTRL_KONDITIONEN", 80)
                TableControl.verticalScrollbar.position = TableControl.verticalScrollbar.Maximum
                TableControl = Session.FindByNameEx("SAPLV69ATCTRL_KONDITIONEN", 80)
                TableRow = TableControl.Rows(0)
                TableRow.Item(1).Text = "ZFAK"
                TableRow.Item(3).Text = "0.01"
                TableRow.Item(4).Text = "GTQ"
                Session.SendVKey(0)
            Next
            Session.FindById("wnd[0]/tbar[0]/btn[11]").Press()
            If Not Session.FindById("wnd[1]/usr/btnSPOP-VAROPTION1") Is Nothing Then
                Session.FindById("wnd[1]/usr/btnSPOP-VAROPTION1").Press()
            End If
            RDT.Rows.Add(New Object() {DR("PURCH#DOC#"), Session.StatusBarText})
        Next

    End Sub

End Class

Public Class MEK2

    Private Session As SAPGUI = Nothing

    Public Sub New(ByVal Box As String)

        Session = New SAPGUI(Box)

    End Sub

    Public Sub Execute(ByVal DT As DataTable, ByVal Condition As String)

        If Not Session.LoggedIn Then
            Exit Sub
        End If

        Dim QVendors = From DR In DT Where DR("Condition") = Condition And IsNumeric(DR("Value")) Group By Code = DR("Vendor") Into G = Group
        Dim VC As String
        Dim TCRI As Integer
        Dim TableControl
        Dim TableRow
        Dim FR() As DataRow

        For Each Vendor In QVendors
            VC = Vendor.Code
            Session.StartTransaction("MEK2")
            Session.FindByNameEx("RV13A-KSCHL", 32).Text = Condition
            Session.SendVKey(0)
            Session.FindById("wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[1,0]").Selected = True
            Session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
            If Condition = "ZDT1" Or Condition = "ZDT4" Or Condition = "ZPVM" Then
                Session.FindByNameEx("F001", 32).Text = "811"
            Else
                Session.FindByNameEx("F001", 32).Text = "2167"
            End If
            Session.FindByNameEx("F002", 32).Text = VC
            Session.FindByNameEx("F003-LOW", 32).Text = ""
            Session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
            TableControl = Session.FindByNameEx("SAPMV13ATCTRL_FAST_ENTRY", 80)
            If TableControl Is Nothing Then Exit Sub
            TCRI = 0
            Do While TableControl.Rows(TCRI).Item(7).Text <> ""
                TableRow = TableControl.Rows(TCRI)
                FR = DT.Select("GCas = '" & TableRow.Item(0).Text & "' And Vendor = '" & VC & "' And Condition = '" & Condition & "'")
                If FR.Count > 0 Then
                    TableRow.Item(2).Text = FR(0)("Value")
                    FR(0).Delete()
                End If
                TCRI += 1
                If TCRI = TableControl.VisibleRowCount Then
                    Session.FindById("wnd[0]").sendVKey(82)
                    TableControl = Session.FindByNameEx("SAPMV13ATCTRL_FAST_ENTRY", 80)
                    TCRI = 1
                End If
            Loop
            DT.AcceptChanges()
            Dim QMaterials = From DR In DT Where DR("Vendor") = VC And DR("Condition") = Condition
            If Not QMaterials Is Nothing Then
                For Each DR In QMaterials
                    TableRow = TableControl.Rows(TCRI)
                    TableRow.Item(0).Text = DR("GCas")
                    TableRow.Item(2).Text = DR("Value")
                    If Not DBNull.Value.Equals(DR("Currency")) Then TableRow.Item(3).Text = DR("Currency")
                    TCRI += 1
                    If TCRI = TableControl.VisibleRowCount Then
                        Session.FindById("wnd[0]").sendVKey(0)
                        Session.FindById("wnd[0]").sendVKey(0)
                        Session.FindById("wnd[0]").sendVKey(82)
                        TableControl = Session.FindByNameEx("SAPMV13ATCTRL_FAST_ENTRY", 80)
                        TCRI = 1
                        If TableControl.Rows(TCRI).Item(7).Text <> "" Then
                            Do While TableControl.Rows(TCRI).Item(7).Text <> ""
                                TCRI += 1
                                If TCRI = TableControl.VisibleRowCount Then Exit For
                            Loop
                        End If
                    End If
                Next
            End If
            Session.SendVKey(0)
            Session.FindById("wnd[0]/tbar[0]/btn[11]").Press()
        Next

    End Sub

    Public Sub Close()

        If Not Session Is Nothing Then
            Session.Close()
        End If

    End Sub

End Class

Public Class OA_Item_New_Validity

    Private Session As SAPGUI = Nothing
    Private SM As String = Nothing
    Private SF As Boolean = False
    Private RF As Boolean = False

    Public ReadOnly Property IsReady() As Boolean

        Get
            IsReady = RF
        End Get

    End Property

    Public ReadOnly Property Success() As Boolean

        Get
            Success = SF
        End Get

    End Property

    Public ReadOnly Property Status() As String

        Get
            Status = SM
        End Get

    End Property

    Public Sub New(ByVal Box As String, ByVal User As String, ByVal Password As String, ByVal OA As String)

        Session = New SAPGUI(Box, User, Password)
        If Not Session.LoggedIn Then
            SM = "SAP Login Failed!"
            Exit Sub
        End If

        Session.StartTransaction("ME32K")
        Session.FindById("wnd[0]/usr/ctxtRM06E-EVRTN").text = OA
        Session.FindById("wnd[0]").sendVKey(0)
        If Session.StatusBarMessageType = "E" Then
            SM = Session.StatusBarText
        Else
            RF = True
        End If

    End Sub

    Public Sub New_Validity(ByVal Item As String, ByVal Price As String)

        If Not RF Then Exit Sub

        SM = Nothing

        If Not IsNumeric(Item) Then
            SM = "Invalid Item Number"
            Exit Sub
        End If

        Item = Val(Item).ToString

        If Not IsNumeric(Price.Trim) Then
            SM = "Invalid Price Format! No update can be made."
            Exit Sub
        End If

        Session.FindById("wnd[0]/usr/txtRM06E-EBELP").text = Item
        Session.FindById("wnd[0]").sendVKey(0)

        If Session.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/txtRM06E-EVRTP[0,0]").text <> Item Then
            SM = "OA Item not found!"
            Exit Sub
        End If

        If Not Session.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/lblRM06E-LOEKZ[13,0]") Is Nothing Then
            SM = "OA Item Blocked or Deleted!"
            Exit Sub
        End If

        Session.FindByNameEx("SAPMM06ETC_0220", 80).Rows(0).Selected = True
        Session.FindById("wnd[0]/tbar[1]/btn[18]").press()
        Session.FindById("wnd[1]/tbar[0]/btn[7]").press()
        Session.FindById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/txtKONP-KBETR[2,0]").text = Math.Round(CDbl(Price.Trim), 2)
        Session.FindById("wnd[0]/tbar[0]/btn[3]").press()

        If Session.StatusBarMessageType = "E" Then
            SM = Session.StatusBarText
            Session.FindById("wnd[0]/tbar[0]/btn[12]").Press()
            Exit Sub
        End If

        If Session.ActiveWindow.Text Like "*Overlapping*" Then
            Session.FindById("wnd[1]/tbar[0]/btn[5]").press()
        End If

        Do While Not Session.StatusBarMessageType = Nothing
            Session.SendVKey(0)
        Loop

    End Sub

    Public Sub Save()

        If Not RF Then Exit Sub

        Session.FindById("wnd[0]/tbar[0]/btn[11]").press()

        If Session.ActiveWindow.Text Like "*Overlapping*" Then
            Session.FindById("wnd[1]/tbar[0]/btn[5]").press()
        End If

        If Not Session.FindById("wnd[1]/usr/btnSPOP-OPTION1") Is Nothing Then
            Session.FindById("wnd[1]/usr/btnSPOP-OPTION1").Press()
        End If

        If Session.StatusBarMessageType = "S" Then
            SF = True
        End If

        SM = Session.StatusBarText
        Session.Close()

    End Sub

End Class

Public Class Credits_ST100_Scripting

    Private Box As String
    Private User As String
    Private Password As String
    Private CI As String
    Private PBk_I As String
    Private Amt_I As String

    Public Sub New(_Box As String, _User As String, _Password As String)

        Box = _Box
        User = _User
        Password = _Password

    End Sub

    Public Sub Execute(Script As Byte, Invoices As List(Of String), Vendor As String, LE As String, Fixed As String, Optional Credits As List(Of String) = Nothing, Optional Main As String = Nothing)

        Select Case Script
            Case 1
                Link(Invoices, Credits, Vendor, LE, Fixed, Main)
            Case 2
                Togle_Block(Invoices, Vendor, LE, Fixed)
        End Select

    End Sub

    Private Sub Link(Invoices As List(Of String), Credits As List(Of String), Vendor As String, LE As String, Fixed As String, Main As String)

        Dim Session As New SAPGUI(Box, User, Password)
        Dim Invoice_Link As String = Nothing
        Dim BkFlag As String = Nothing

        If Not Session.LoggedIn Then
            MsgBox("Failed to connect to SAP - " & Box, MsgBoxStyle.Critical, "Link Script")
            Exit Sub
        End If

        If Credits.Count > 1 Or Invoices.Count > 1 Then
            Session.StartTransaction("XK03")
            Session.FindByNameEx("RF02K-LIFNR", 32).Text = Vendor
            Session.FindByNameEx("RF02K-BUKRS", 32).Text = LE
            Session.FindByNameEx("RF02K-D0215", 42).Selected = True
            Session.SendVKey(0)
            If Session.FindByNameEx("LFB1-XPORE", 42).Selected Then
                MsgBox("This vendor is flagged for Individual Payments! Process Aborted.", MsgBoxStyle.Critical, "Vendor Check")
                Session.Close()
                Exit Sub
            End If
        End If

        Session.StartTransaction("FBL1")
        Session.FindByNameEx("KD_LIFNR-LOW", 32).Text = Vendor
        Session.FindByNameEx("KD_BUKRS-LOW", 32).Text = LE
        Session.FindByNameEx("PA_VARI", 32).Text = "/LINKING"
        Session.FindById("wnd[0]/tbar[1]/btn[8]").Press()

        PBk_I = Session.FindByText("PBk").ID.ToString
        PBk_I = PBk_I.Substring(PBk_I.LastIndexOf("[") + 1, PBk_I.LastIndexOf(",") - PBk_I.LastIndexOf("[") - 1)

        Amt_I = Session.FindByText("Amount in doc. curr.").ID.ToString
        Amt_I = Amt_I.Substring(Amt_I.LastIndexOf("[") + 1, Amt_I.LastIndexOf(",") - Amt_I.LastIndexOf("[") - 1)

        For Each Invoice As String In Invoices
            CI = GUI_Find(Session, Invoice)
            If CI Is Nothing Then
                MsgBox("Invoice " & Invoice & " could not be found! Process Aborted.", MsgBoxStyle.Critical, "Invoices Verification")
                Session.Close()
                Exit Sub
            End If
            BkFlag = Session.FindById("wnd[0]/usr/lbl[*,%]".Replace("*", PBk_I).Replace("%", CI)).Text
            If BkFlag <> "" And BkFlag <> "C" Then
                MsgBox("Invoice " & Invoice & " is currently blocked! Process Aborted.", MsgBoxStyle.Critical, "Invoices Verification")
                Session.Close()
                Exit Sub
            End If
        Next

        For Each Credit As String In Credits
            CI = GUI_Find(Session, Credit)
            If CI Is Nothing Then
                MsgBox("Credit " & Credit & " could not be found! Process Aborted.", MsgBoxStyle.Critical, "Credits Verification")
                Session.Close()
                Exit Sub
            End If
            BkFlag = Session.FindById("wnd[0]/usr/lbl[*,%]".Replace("*", PBk_I).Replace("%", CI)).Text
            If BkFlag <> "" And BkFlag <> "C" And LE <> "273" And LE <> "682" Then
                MsgBox("Credit " & Credit & " is currently blocked! Process Aborted.", MsgBoxStyle.Critical, "Credits Verification")
                Session.Close()
                Exit Sub
            End If
        Next

        If Invoices.Count > 1 Then
            Invoice_Link = Main
            For Each Invoice As String In Invoices
                If Invoice <> Invoice_Link Then
                    CI = GUI_Find(Session, Invoice)
                    Session.SendVKey(2)
                    Session.FindById("wnd[0]/tbar[1]/btn[13]").Press()
                    If Session.FindByNameEx("BSEG-ZBFIX", 32).changeable Then Session.FindByNameEx("BSEG-ZBFIX", 32).Text = Fixed
                    Session.FindByNameEx("BSEG-REBZG", 31).Text = Invoice_Link
                    If Session.FindByNameEx("BSEG-SGTXT", 32).Text.ToString.Length < 45 Then Session.FindByNameEx("BSEG-SGTXT", 32).Text = Session.FindByNameEx("BSEG-SGTXT", 32).Text & " /02DB"
                    Session.SendVKey(0)
                    Session.FindById("wnd[0]/tbar[0]/btn[11]").Press() '**** --> Save <--- *****
                    Session.SendVKey(0)
                End If
            Next
        Else
            Invoice_Link = Invoices(0)
        End If

        For Each Credit As String In Credits
            CI = GUI_Find(Session, Credit)
            Session.SendVKey(2)
            Session.FindById("wnd[0]/tbar[1]/btn[13]").Press()
            If Session.FindByNameEx("BSEG-ZBFIX", 32).changeable Then Session.FindByNameEx("BSEG-ZBFIX", 32).Text = Fixed
            Session.FindByNameEx("BSEG-REBZG", 31).Text = Invoice_Link
            If Session.FindByNameEx("BSEG-SGTXT", 32).Text.ToString.Length < 45 Then Session.FindByNameEx("BSEG-SGTXT", 32).Text = Session.FindByNameEx("BSEG-SGTXT", 32).Text & " /02DB"
            Session.SendVKey(0)
            Session.FindById("wnd[0]/tbar[0]/btn[11]").Press() '**** --> Save <--- *****
            Session.SendVKey(0)
        Next

        Session.Close()

    End Sub

    Private Sub Togle_Block(Documents As List(Of String), Vendor As String, LE As String, Fixed As String)

        Dim Session As New SAPGUI(Box, User, Password)
        Dim BkFlag As String = Nothing

        If Not Session.LoggedIn Then
            MsgBox("Failed to connect to SAP - " & Box, MsgBoxStyle.Critical, "Block Script")
            Exit Sub
        End If

        Session.StartTransaction("FBL1")
        Session.FindByNameEx("KD_LIFNR-LOW", 32).Text = Vendor
        Session.FindByNameEx("KD_BUKRS-LOW", 32).Text = LE
        Session.FindByNameEx("PA_VARI", 32).Text = "/LINKING"
        Session.FindById("wnd[0]/tbar[1]/btn[8]").Press()

        PBk_I = Session.FindByText("PBk").ID.ToString
        PBk_I = PBk_I.Substring(PBk_I.LastIndexOf("[") + 1, PBk_I.LastIndexOf(",") - PBk_I.LastIndexOf("[") - 1)

        For Each Document As String In Documents
            CI = GUI_Find(Session, Document)
            BkFlag = Session.FindById("wnd[0]/usr/lbl[*,%]".Replace("*", PBk_I).Replace("%", CI)).Text
            If BkFlag <> "" And BkFlag <> "C" Then
                MsgBox("Block type " & BkFlag & " on Document " & Document & " is not allowed be modified.", MsgBoxStyle.Critical, "Block Script")
            Else
                Session.SendVKey(2)
                Session.FindById("wnd[0]/tbar[1]/btn[13]").Press()
                If Session.FindByNameEx("BSEG-ZLSPR", 32).Text <> "" Then
                    Session.FindByNameEx("BSEG-ZLSPR", 32).Text = ""
                Else
                    Session.FindByNameEx("BSEG-ZLSPR", 32).Text = "C"
                    If Session.FindByNameEx("BSEG-SGTXT", 32).Text.ToString.Length + (" Block Request by: " & Fixed).Length <= 50 Then
                        Session.FindByNameEx("BSEG-SGTXT", 32).Text = Session.FindByNameEx("BSEG-SGTXT", 32).Text & " Block Request by: " & Fixed
                    End If
                End If
                Session.SendVKey(0)
                Session.FindById("wnd[0]/tbar[0]/btn[11]").Press() '**** --> Save <--- *****
                Session.SendVKey(0)
            End If
        Next

        Session.Close()

    End Sub

    Private Function GUI_Find(Session As SAPGUI, Reference As String) As String

        GUI_Find = Nothing

        Session.SendVKey(71)
        Session.FindById("wnd[1]/usr/chkSCAN_STRING-START").selected = False
        Session.FindById("wnd[1]/usr/txtRSYSF-STRING").text = Reference
        Session.FindById("wnd[1]/tbar[0]/btn[0]").press()
        If Session.FindById("wnd[2]").text = "Information" Then
            Session.FindById("wnd[2]/tbar[0]/btn[0]").press()
            Session.FindById("wnd[1]/tbar[0]/btn[12]").press()
            Exit Function
        End If

        Dim I As Integer = 2
        Dim FI As Integer = 0
        Do While Not Session.FindById("wnd[2]/usr/lbl[9," & I & "]") Is Nothing
            If Session.FindById("wnd[2]/usr/lbl[9," & I & "]").Text = Reference Then
                FI = I
                Exit Do
            End If
            I += 1
        Loop

        If FI = 0 Then Exit Function

        Session.FindById("wnd[2]/usr/lbl[9," & FI & "]").setFocus()
        Session.FindById("wnd[2]").sendVKey(2)

        'Session.FindById("wnd[2]/usr").horizontalScrollbar.position = 127
        'If Session.FindById("wnd[2]/usr/lbl[7,2]") Is Nothing OrElse Session.FindById("wnd[2]/usr/lbl[7,2]").Text <> Reference Then
        '    Exit Function
        'End If
        'Session.FindById("wnd[2]/usr/lbl[7,2]").setFocus()
        'Session.FindById("wnd[2]").sendVKey(2)

        GUI_Find = Session.ActiveWindow.SystemFocus.ID.Substring(Session.ActiveWindow.SystemFocus.ID.IndexOf(",") + 1).TrimEnd("]")

    End Function

End Class

Public Class AP_Trade

    Private Session As SAPGUI = Nothing
    Private LE() As String
    Private Vendor() As String

    Public Sub New(Box As String, User As String, Password As String)

        Session = New SAPGUI(Box, User, Password)

    End Sub

    Public Sub New(Box As String)

        Session = New SAPGUI(Box)

    End Sub

    Public Function Get_Data(LE() As String, Vendors() As String) As DataTable

        Get_Data = Nothing
        If Not Session.LoggedIn Then
            Exit Function
        End If

        Session.StartTransaction("F.98")
        Session.FindById("wnd[0]/usr/lbl[5,6]").setFocus()
        Session.FindById("wnd[0]").sendVKey(2)
        Session.FindById("wnd[0]/usr/lbl[12,15]").setFocus()
        Session.FindById("wnd[0]").sendVKey(2)

        Session.FindByNameEx("P_REGION", 32).Text = "LA"

        Session.ArrayToClipboard(LE)
        Session.FindByNameEx("%_S_BUKRS_%_APP_%-VALU_PUSH", 40).Press()
        Session.FindById("wnd[1]/tbar[0]/btn[16]").Press()
        Session.FindById("wnd[1]/tbar[0]/btn[24]").Press()
        Session.FindById("wnd[1]/tbar[0]/btn[8]").Press()

        Session.FindByNameEx("%_S_KTOKK_%_APP_%-VALU_PUSH", 40).Press()
        Session.FindById("wnd[1]/tbar[0]/btn[16]").Press()
        Session.FindById("wnd[1]/tbar[0]/btn[8]").Press()

        Session.ArrayToClipboard(Vendors)
        Session.FindByNameEx("%_S_LIFNR_%_APP_%-VALU_PUSH", 40).Press()
        Session.FindById("wnd[1]/tbar[0]/btn[16]").Press()
        Session.FindById("wnd[1]/tbar[0]/btn[24]").Press()
        Session.FindById("wnd[1]/tbar[0]/btn[8]").Press()

        Session.FindById("wnd[0]/tbar[1]/btn[8]").Press()

        Session.FindById("wnd[0]/tbar[0]/okcd").text = "%pc"
        Session.FindById("wnd[0]").sendVKey(0)

        Session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Selected = True
        Session.FindById("wnd[1]/tbar[0]/btn[0]").Press()

        Session.FindById("wnd[1]/usr/ctxtDY_PATH").text = My.Computer.FileSystem.SpecialDirectories.Temp & "\"
        Session.FindById("wnd[1]/usr/ctxtDY_FILENAME").text = "APTrade.txt"
        Session.FindById("wnd[1]/tbar[0]/btn[11]").Press()

        Session.Close()

        Get_Data = TD_TextFile_Data(My.Computer.FileSystem.SpecialDirectories.Temp & "\APTrade.txt")
        If Not Get_Data.Columns("Column1") Is Nothing Then
            Get_Data.Columns.Remove("Column1")
        End If

    End Function

    Private Function TD_TextFile_Data(ByVal Path As String, Optional ByVal SkipTopLines As Integer = 0, Optional ByVal NoHeaders As Boolean = False) As Object

        TD_TextFile_Data = Nothing
        Try
            Dim F As New Microsoft.VisualBasic.FileIO.TextFieldParser(Path)
            F.TextFieldType = FileIO.FieldType.Delimited
            F.SetDelimiters(Chr(9))

            Dim R As String() = Nothing
            Dim D As New DataTable
            Dim CI As Integer
            Dim I As Integer = 1

            Do While Not F.EndOfData And I <= SkipTopLines
                R = F.ReadFields
                I += 1
            Loop

            If NoHeaders Then
                For I = 1 To R.Count
                    D.Columns.Add("Field" & I, Type.GetType("System.String"))
                Next
            Else
                If Not F.EndOfData Then
                    R = F.ReadFields
                    If R.Length > 0 Then
                        CI = 1
                        For Each CN As String In R
                            Do While Not D.Columns(CN) Is Nothing
                                CN = CN & CI
                                CI += 1
                            Loop
                            D.Columns.Add(CN, Type.GetType("System.String"))
                        Next
                    End If
                End If
            End If

            While Not F.EndOfData
                Try
                    R = F.ReadFields
                    D.LoadDataRow(R, True)
                Catch ex As Exception
                End Try
            End While

            If D.Rows.Count > 0 Then TD_TextFile_Data = D
            F.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

End Class