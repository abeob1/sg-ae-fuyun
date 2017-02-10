Imports System.Xml

Module modMarketingDocuments

    Public Function BP_ADD(ByVal oDICompany As SAPbobsCOM.Company, _
                           ByVal oDITargetComp As SAPbobsCOM.Company, _
                           ByVal sCardCode As String, _
                           ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   BP_ADD()
        '   Purpose     :   This function will create Business Partner Master in Target Database
        '               
        '   Parameters  :   Byval oCompany As SAPbobsCOM.Company
        '                      oCompany =  set the SAP DI Company Object
        '                   Byval sCardCode as String
        '                       sCardCode = Cardcode
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   SRI
        '   Date        :   21 April 2011
        '   Change      :   
        ' ************************************************************************************

        Dim sFuncName As String = String.Empty
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oBPTarget As SAPbobsCOM.BusinessPartners
        Dim sXMLFile As String = String.Empty
        Dim lRetCode As Long
        Dim iTemp As Integer
        Dim oRS As SAPbobsCOM.Recordset
        Dim sSql As String = String.Empty

        Try
            sFuncName = "BP_ADD()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            oBP = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            oBPTarget = oDITargetComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

            sXMLFile = System.IO.Directory.GetCurrentDirectory & "\" & Now.Date.ToString("yyyyMMdd") & Now.ToString("HHmmss") & sCardCode & ".xml"

            oRS = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Check BP Exists in Master DB", sFuncName)

            If oBP.GetByKey(sCardCode) Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Saving BP as xml file.", sFuncName)

                oDICompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                oBP.SaveXML(sXMLFile)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Update_XMLFile()", sFuncName)
                Update_XMLFile(sXMLFile, oDICompany, oDITargetComp)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating BP " & sCardCode & " in Target DB : " & oDITargetComp.CompanyDB.ToString, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetBusinessObjectFromXML", sFuncName)
                oBPTarget = oDITargetComp.GetBusinessObjectFromXML(sXMLFile, 0)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling oBPTarget.Add()", sFuncName)
                lRetCode = oBPTarget.Add()
                If lRetCode <> 0 Then
                    oDITargetComp.GetLastError(iTemp, sErrDesc)
                    Throw New ArgumentException(iTemp & " : " & sErrDesc)
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successflly Created BP" & sCardCode & " in Target DB : " & oDITargetComp.CompanyDB.ToString, sFuncName)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            BP_ADD = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            If System.IO.File.Exists(sXMLFile) Then
                System.IO.File.Delete(sXMLFile)
            End If
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            BP_ADD = RTN_ERROR
        Finally
            If System.IO.File.Exists(sXMLFile) Then
                System.IO.File.Delete(sXMLFile)
            End If
        End Try

    End Function

    Public Function BP_UPDATE(ByVal oDICompany As SAPbobsCOM.Company, ByVal oDITargetComp As SAPbobsCOM.Company, _
                              ByVal sCardCode As String, ByVal sTargetCardCode As String, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   BP_UPDATE()
        '   Purpose     :   This function will Update Business Partner Master in Target Database
        '               
        '   Parameters  :   Byval oCompany As SAPbobsCOM.Company
        '                      oCompany =  set the SAP DI Company Object
        '                   Byval sCardCode as String
        '                       sCardCode = Cardcode
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   SRI
        '   Date        :   21 April 2011
        '   Change      :   
        ' ********************************************************************************** 

        Dim sFuncName As String = String.Empty
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oBPTarget As SAPbobsCOM.BusinessPartners
        Dim sXMLFile As String = String.Empty
        Dim lRetCode As Long
        Dim iTemp As Integer
        Dim oRS As SAPbobsCOM.Recordset
        Dim sSql As String = String.Empty

        Try
            sFuncName = "BP_UPDATE()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            oBP = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            oBPTarget = oDITargetComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

            oRS = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sXMLFile = System.IO.Directory.GetCurrentDirectory & "\" & Now.Date.ToString("yyyyMMdd") & Now.ToString("HHmmss") & sCardCode & ".xml"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Check BP Exists in Master DB", sFuncName)

            If oBP.GetByKey(sCardCode) Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Saving BP as xml file.", sFuncName)
                oBP.SaveXML(sXMLFile)
                oDITargetComp.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Update_XMLFile()", sFuncName)
                Update_XMLFile(sXMLFile, oDICompany, oDITargetComp)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Check BP Exists in Target DB :" & oDITargetComp.CompanyDB.ToString, sFuncName)
                If oBPTarget.GetByKey(sTargetCardCode) Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating BP " & sCardCode & " in Target DB : " & oDITargetComp.CompanyDB.ToString, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetBusinessObjectFromXML", sFuncName)
                    oBPTarget = oDITargetComp.GetBusinessObjectFromXML(sXMLFile, 0)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling oBPTarget.Update()", sFuncName)

                    lRetCode = oBPTarget.Update()
                    If lRetCode <> 0 Then
                        oDITargetComp.GetLastError(iTemp, sErrDesc)
                        Throw New ArgumentException(iTemp & " : " & sErrDesc)
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successflly Updated BP" & sCardCode & " in Target DB : " & oDITargetComp.CompanyDB.ToString, sFuncName)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("BP" & sCardCode & " doesn't exists in Target DB : " & oDITargetComp.CompanyDB.ToString, sFuncName)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            BP_UPDATE = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            If System.IO.File.Exists(sXMLFile) Then
                System.IO.File.Delete(sXMLFile)
            End If
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            BP_UPDATE = RTN_ERROR
        Finally
            If System.IO.File.Exists(sXMLFile) Then
                System.IO.File.Delete(sXMLFile)
            End If
        End Try

    End Function

    Public Function SyncBPMaster(ByVal sCardCode As String, ByVal oDICompany As SAPbobsCOM.Company, ByVal oDITargetComp As SAPbobsCOM.Company, ByRef sErrDesc As String)

        ' **********************************************************************************
        '   Function    :   SyncBPMaster()
        '   Purpose     :   This function will Sycn Business Partner Master Data 
        '               
        '   Parameters  :   Byval sCardCode as String
        '                       sCardCode=BP Code
        '                   Byval oDICompany As SAPbobsCOM.Company
        '                      oDICompany =  set the SAP DI Company Object        '
        '                   Byval oDITargetComp As SAPbobsCOM.Company
        '                      oDITargetComp =  set the SAP DI Company Object
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   SRI
        '   Date        :   30 April 2013
        '   Change      :   
        ' ********************************************************************************** 

        Dim sFuncName As String = String.Empty
        Dim oDs As New DataSet
        Dim sSQL As String = String.Empty
        Dim oRs As SAPbobsCOM.Recordset
        Dim oRsTarget As SAPbobsCOM.Recordset
        Dim bIsError As Boolean = False
        Dim sSyncBPCode As String = String.Empty
        Dim sTargetCardCode As String = String.Empty

        Try
            sFuncName = "SyncBPMaster()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            oRs = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsTarget = oDITargetComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


            sSQL = "SELECT U_SYNCCODE FROM OCRD WHERE CARDCODE='" & sCardCode & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSQL, sFuncName)
            oRs.DoQuery(sSQL)

            sSyncBPCode = oRs.Fields.Item(0).Value.ToString

            sSQL = "SELECT CARDCODE FROM OCRD WHERE U_SYNCCODE='" & sSyncBPCode & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSQL, sFuncName)
            oRsTarget.DoQuery(sSQL)

            sTargetCardCode = oRsTarget.Fields.Item(0).Value.ToString

            If oRsTarget.EoF Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling BP_ADD", sFuncName)
                If BP_ADD(oDICompany, oDITargetComp, sCardCode, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling BP_UPDATE", sFuncName)
                If BP_UPDATE(oDICompany, oDITargetComp, sCardCode, sTargetCardCode, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            SyncBPMaster = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            SyncBPMaster = RTN_ERROR
        Finally

        End Try
    End Function

    Public Sub Update_XMLFile(ByRef sXMLFile As String, _
                              ByVal oDICompany As SAPbobsCOM.Company, ByVal oDITargetComp As SAPbobsCOM.Company)

        Dim oXMLDoc As New XmlDocument()
        Dim oNode As XmlNode
        Dim sFuncName As String = String.Empty
        Dim oNodelist As XmlNodeList
        Dim sAcctCode As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oRs As SAPbobsCOM.Recordset

        Try
            sFuncName = "Update_XMLFile"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function..", sFuncName)
            oXMLDoc.Load(sXMLFile)

            oRs = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ' Change Usersign
            oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OCRD/row/UserSign")
            oNode.InnerText = oDITargetComp.UserSignature
            oXMLDoc.Save(sXMLFile)

            oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OCRD/row/AutoPost")
            If Not IsNothing(oNode) Then
                oNode.InnerText = "N"
                oXMLDoc.Save(sXMLFile)
            End If

            oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OCRD/row/IndustryC")
            If Not IsNothing(oNode) Then
                oNode.ParentNode.RemoveChild(oNode)
                oXMLDoc.Save(sXMLFile)
            End If

            oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OCRD/row/DflAgrmnt")
            If Not IsNothing(oNode) Then
                oNode.ParentNode.RemoveChild(oNode)
                oXMLDoc.Save(sXMLFile)
            End If

            ' Delete Balances
            oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OCRD/row/Balance")
            If Not IsNothing(oNode) Then
                oNode.ParentNode.RemoveChild(oNode)
                oXMLDoc.Save(sXMLFile)
            End If

            oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OCRD/row/ChecksBal")
            If Not IsNothing(oNode) Then
                oNode.ParentNode.RemoveChild(oNode)
                oXMLDoc.Save(sXMLFile)
            End If

            oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OCRD/row/DNotesBal")
            If Not IsNothing(oNode) Then
                oNode.RemoveAll()
                oXMLDoc.Save(sXMLFile)
            End If

            oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OCRD/row/OrdersBal")
            If Not IsNothing(oNode) Then
                oNode.RemoveAll()
                oXMLDoc.Save(sXMLFile)
            End If

            oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OCRD/row/AbsEntry")
            If Not IsNothing(oNode) Then
                oNode.RemoveAll()
                oXMLDoc.Save(sXMLFile)
            End If

            '1. GET Groupcode from TargetDB based on Master DB GroupName and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OCRD/row/GroupCode", "OCRG", "GroupCode", "GroupName", True, oXMLDoc, sXMLFile)

            '2. GET ShipType code from TargetDB based on Master DB ShipTypeName and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OCRD/row/ShipType", "OSHP", "TrnspCode", "TrnspName", True, oXMLDoc, sXMLFile)

            '3. GET Industry code from TargetDB based on Master DB IndustryName and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OCRD/row/Industry", "OOND", "IndCode", "IndName", True, oXMLDoc, sXMLFile)

            '4. GET SalesEmployee code from TargetDB based on Master DB SalesEmployee Name and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OCRD/row/SlpCode", "OSLP", "SlpCode", "SlpName", True, oXMLDoc, sXMLFile)

            '5. GET Paymentterms code from TargetDB based on Master DB Paymentterms Name and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OCRD/row/GroupNum", "OCTG", "GroupNum", "PymntGroup", True, oXMLDoc, sXMLFile)

            '6. GET Creditcard from TargetDB based on Master DB card Name and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OCRD/row/CreditCard", "OCRC", "CreditCard", "CardName", True, oXMLDoc, sXMLFile)

            '7. GET PriceListNum from TargetDB based on Master DB PriceList Name and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OCRD/row/ListNum", "OPLN", "ListNum", "ListName", True, oXMLDoc, sXMLFile)

            '8. GET LangCode from TargetDB based on Master DB Lang Name and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OCRD/row/LangCode", "OLNG", "Code", "ShortName", True, oXMLDoc, sXMLFile)

            '9. GET DebPayAcct from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OCRD/row/DebPayAcct", "OACT", "AcctCode", "FormatCode", False, oXMLDoc, sXMLFile)

            '10. GET Downpayment Clearing Account  from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OCRD/row/DpmClear", "OACT", "AcctCode", "FormatCode", False, oXMLDoc, sXMLFile)

            '11. GET Downpayment Interim Account from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OCRD/row/DpmIntAct", "OACT", "AcctCode", "FormatCode", False, oXMLDoc, sXMLFile)

            '12. GET Technician  from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OCRD/row/DfTcnician", "OHEM", "empID", "firstName+lastName", False, oXMLDoc, sXMLFile)

            '13. GET Territory  from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OCRD/row/Territory", "OTER", "territryID", "descript", False, oXMLDoc, sXMLFile)

            '14. GET Control Accounts  from TargetDB based on Master DB and Update 
            oNodelist = oXMLDoc.SelectNodes("/BOM/BO/CRD3/row/AcctCode")
            If Not IsNothing(oNodelist) Then
                For Each oNode In oNodelist
                    sAcctCode = oNode.InnerText
                    sSQL = "SELECT AcctCode from  [" & oDITargetComp.CompanyDB.ToString & "].[dbo].OACT WHERE FormatCode in (select FormatCode from OACT WHERE AcctCode='" & sAcctCode & "')"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL Query" & sSQL, sFuncName)
                    oRs.DoQuery(sSQL)
                    If Not oRs.EoF Then oNode.InnerText = oRs.Fields.Item("AcctCode").Value
                    oXMLDoc.Save(sXMLFile)
                Next
            End If

            oNode = Nothing
            oXMLDoc = Nothing

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch ex As Exception

        End Try

    End Sub


End Module
