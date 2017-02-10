Module modSync

    Public Function CopyDataToOtherEntity(ByRef oForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        'Function   :   CopyDataToOtherEntity()
        'Purpose    :   To Copy Data to Other Entities
        'Parameters :   ByVal oForm As SAPbouiCOM.Form
        '                   oForm=Form Type
        '               ByRef sErrDesc as String
        '                   Returns Error description
        'Return     :   0 - FAILURE
        '               1 - SUCCESS
        'Author     :   Sri
        'Date       :   30 April 2013
        'Change     :
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim oDITargetComp As SAPbobsCOM.Company = Nothing
        Dim sDBName As String
        Dim sSAPUser As String = String.Empty
        Dim sSAPPWD As String = String.Empty
        Dim iCnt As Integer
        Dim oCheckBoxDB As SAPbouiCOM.CheckBox
        Dim bIsChecked As Boolean = False
        Dim sCardCode As String = String.Empty
        Dim sItemCode As String = String.Empty
        Dim oStaticItem As SAPbouiCOM.StaticText
        Dim oStaticCode As SAPbouiCOM.StaticText
        Try
            sFuncName = "CopyDataToOtherEntity()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DisplayStatus()", sFuncName)
            Call DisplayStatus(oForm, "Please wait while processing ......", sErrDesc)

            ' Validation for DBMatrix
            oMatrix = oForm.Items.Item("DBMatrix").Specific
            If oMatrix.RowCount = 0 Then
                p_oSBOApplication.MessageBox("No Records in the Item Matrix...")
                Exit Function
            End If

            'The Checkbox click or not Validation for DBMatrix 
            For iCnt = 1 To oMatrix.RowCount ' Start DBMatrix loop
                oCheckBoxDB = oMatrix.Columns.Item("DBCheck").Cells.Item(iCnt).Specific
                If oCheckBoxDB.Checked = True Then
                    bIsChecked = True
                    Exit For
                End If
            Next

            If bIsChecked = False Then
                p_oSBOApplication.MessageBox("Please select DataBase Name....")
                Exit Function
            End If

            ' Get DataBase Value from Matrix2
            For iCnt = 1 To oMatrix.RowCount ' Start DBMatrix loop
                oCheckBoxDB = oMatrix.Columns.Item("DBCheck").Cells.Item(iCnt).Specific

                If oCheckBoxDB.Checked = True Then ' If Checkbox click DBMatrix
                    sDBName = oMatrix.Columns.Item("DBName").Cells.Item(iCnt).Specific.string
                    sSAPUser = oMatrix.Columns.Item("SAPUser").Cells.Item(iCnt).Specific.string
                    sSAPPWD = oMatrix.Columns.Item("SAPPwd").Cells.Item(iCnt).Specific.string

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToCompany()", sFuncName)
                    If ConnectTargetDB(oDITargetComp, sDBName, sSAPUser, sSAPPWD, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                    oStaticItem = oForm.Items.Item("Type").Specific
                    oStaticCode = oForm.Items.Item("Code").Specific

                    Select Case oStaticItem.Caption
                        Case "BP" ' Business Partner
                            p_oSBOApplication.StatusBar.SetText("Copying BP: " & oStaticCode.Caption & " Data to Database: " & sDBName & " .....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SyncBPMaster()", sFuncName)
                            If SyncBPMaster(oStaticCode.Caption, p_oDICompany, oDITargetComp, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If
                            p_oSBOApplication.StatusBar.SetText("Successfully copied BP: " & oStaticCode.Caption & " Data to Database: " & sDBName & ".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                        Case "Item"
                            p_oSBOApplication.StatusBar.SetText("Copying Item: " & oStaticCode.Caption & " Data to Database: " & sDBName & " .....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SyncItemMaster()", sFuncName)
                            If SyncItemMaster(oStaticCode.Caption, p_oDICompany, oDITargetComp, sErrDesc) <> RTN_SUCCESS Then
                                Throw New ArgumentException(sErrDesc)
                            End If
                            p_oSBOApplication.StatusBar.SetText("Successfully copied Item: " & oStaticCode.Caption & " Data to Database: " & sDBName & ".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            oDITargetComp.Disconnect()
                            oDITargetComp = Nothing
                    End Select
                End If
            Next
            p_oSBOApplication.StatusBar.SetText("Successfully Copied Data to Other Entities.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            CopyDataToOtherEntity = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch exc As Exception
            sErrDesc = exc.Message
            CopyDataToOtherEntity = RTN_ERROR
            RollBackTransaction(sErrDesc)
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        Finally
            oMatrix = Nothing
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
            Call EndStatus(sErrDesc)
        End Try

    End Function

    Public Function LoadDataSyncForm(ByVal sCode As String, ByVal sType As String, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        'Function   :   LoadDataSyncForm()
        'Purpose    :   Handle Form Load for Data Sync
        'Parameters :   ByVal oForm As SAPbouiCOM.Form
        '                   oForm=Form Type
        'Return     :   0 - FAILURE
        '               1 - SUCCESS
        'Author     :   Sri
        'Date       :   30 April 2013
        'Change     :
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim oForm As SAPbouiCOM.Form
        Dim oDBMatrix As SAPbouiCOM.Matrix
        Dim oCol As SAPbouiCOM.Column

        Try
            sFuncName = "LoadDataSyncForm()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            LoadFormFrmXml("AB_DataSync", "AB_DataSync.srf", False)
            oForm = p_oSBOApplication.Forms.ActiveForm

            oForm.Freeze(True)
            'Add User DAtasource For Matrix
            AddUserDataSrc(oForm, "DBCheck", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_TEXT, 20)
            AddUserDataSrc(oForm, "DBName", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_TEXT, 50)
            AddUserDataSrc(oForm, "SAPUser", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_TEXT, 50)
            AddUserDataSrc(oForm, "SAPPwd", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_TEXT, 50)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Set Matrix: DBMatrix", sFuncName)
            oDBMatrix = oForm.Items.Item("DBMatrix").Specific
            oDBMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            oCol = oDBMatrix.Columns.Item("DBName")
            oCol.DataBind.SetBound(True, , "DBName")

            oCol = oDBMatrix.Columns.Item("SAPUser")
            oCol.DataBind.SetBound(True, , "SAPUser")

            oCol = oDBMatrix.Columns.Item("SAPPwd")
            oCol.DataBind.SetBound(True, , "SAPPwd")

            oDBMatrix.Columns.Item("SAPUser").Visible = False
            oDBMatrix.Columns.Item("SAPPwd").Visible = False
            oDBMatrix.Columns.Item("DBName").Width = 200

            oForm.Visible = True
            oDBMatrix.AutoResizeColumns()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling LoadDataToMatrix()", sFuncName)
            If LoadDataToMatrix(sCode, sType, oForm, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            oForm.Freeze(False)
            LoadDataSyncForm = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch exc As Exception
            LoadDataSyncForm = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        Finally

        End Try

    End Function

    Public Function LoadDataToMatrix(ByVal sCode As String, ByVal sType As String, ByVal oForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        'Function   :   LoadDataToMatrix()
        'Purpose    :   Handle Form Load for Data Sync
        'Parameters :   ByVal oForm As SAPbouiCOM.Form
        '                   oForm=Form Type
        'Return     :   0 - FAILURE
        '               1 - SUCCESS
        'Author     :   Sri
        'Date       :   30 April 2013
        'Change     :
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oDataTable As SAPbouiCOM.DataTable
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim oStaticItem As SAPbouiCOM.StaticText

        Try
            sFuncName = "LoadDataToMatrix()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sSQL = "select * from [@AB_DBNAMES]"

            If oForm.DataSources.DataTables.Count.Equals(0) Then
                oForm.DataSources.DataTables.Add("DBSync")
            Else
                oForm.DataSources.DataTables.Item("DBSync").Clear()
            End If
            oDataTable = oForm.DataSources.DataTables.Item("DBSync")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL : " & sSQL, sFuncName)
            oDataTable.ExecuteQuery(sSQL)

            oMatrix = oForm.Items.Item("DBMatrix").Specific

            If oDataTable.Rows.Count > 0 Then
                oMatrix.Columns.Item("DBName").DataBind.Bind("DBSync", "Name")
                oMatrix.Columns.Item("SAPUser").DataBind.Bind("DBSync", "U_SAPUSER")
                oMatrix.Columns.Item("SAPPwd").DataBind.Bind("DBSync", "U_SAPPWD")
                oMatrix.LoadFromDataSource()
            End If

            oStaticItem = oForm.Items.Item("Type").Specific
            oStaticItem.Caption = sType
            oStaticItem = oForm.Items.Item("Code").Specific
            oStaticItem.Caption = sCode

            oDataTable = Nothing


            LoadDataToMatrix = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch exc As Exception
            LoadDataToMatrix = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        Finally

        End Try
    End Function

    Public Function LoadFormFrmXml(ByVal formType As String, ByVal xmlFile As String, ByVal isModal As Boolean) As SAPbouiCOM.Form

        Dim fcp As SAPbouiCOM.FormCreationParams
        Dim Form As SAPbouiCOM.Form
        Dim UniqueId As String

        UniqueId = formType

        fcp = p_oSBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
        fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
        fcp.FormType = formType
        fcp.UniqueID = UniqueId
        fcp.XmlData = LoadFromXML(xmlFile)
        Form = p_oSBOApplication.Forms.AddEx(fcp)

        Return Form
    End Function

    Private Function LoadFromXML(ByRef fileName As String) As String
        Dim XmlDoc As Xml.XmlDocument
        Dim Path As String
        XmlDoc = New Xml.XmlDocument
        '// load the content of the XML File
        Path = System.Windows.Forms.Application.StartupPath
        If Path.EndsWith("bin") Then
            'Path = Path.Remove(Path.Length - 3, 3)
            Path += "\"
        Else
            Path += "\Forms\"
            'Path += "\"
        End If
        XmlDoc.Load(Path & fileName)
        '// load the form to the SBO application in one batch
        Return (XmlDoc.InnerXml)
    End Function

End Module
