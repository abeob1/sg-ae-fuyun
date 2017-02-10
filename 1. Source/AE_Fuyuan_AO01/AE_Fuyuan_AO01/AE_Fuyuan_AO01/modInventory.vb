Imports System.Xml

Module modInventory

#Region "Item Master Sync"

    Public Function ItemMaster_ADD(ByVal oDICompany As SAPbobsCOM.Company, ByVal oDITargetComp As SAPbobsCOM.Company, _
                                   ByVal sItemcode As String, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   ItemMaster_ADD()
        '   Purpose     :   This function will create BOM in Target Database
        '               
        '   Parameters  :   Byval oCompany As SAPbobsCOM.Company
        '                      oCompany =  set the SAP DI Company Object
        '                   Byval sItemcode as String
        '                       sItemcode = Itemcode
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   SRI
        '   Date        :   25 April 2011
        '   Change      :   
        ' ************************************************************************************

        Dim sFuncName As String = String.Empty
        Dim oItems As SAPbobsCOM.Items
        Dim oItemsTarget As SAPbobsCOM.Items
        Dim sXMLFile As String = String.Empty
        Dim lRetCode As Long
        Dim iTemp As Integer

        Try
            sFuncName = "ItemMaster_ADD()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            oItems = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
            oItemsTarget = oDITargetComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

            sXMLFile = System.Windows.Forms.Application.StartupPath & "\Item_" & Now.Date.ToString("yyyyMMdd") & Now.ToString("HHmmss") & ".xml"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Check Item Exists in Master DB", sFuncName)

            If oItems.GetByKey(sItemcode) Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Saving Item Master as xml file.", sFuncName)

                oDICompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                oItems.SaveXML(sXMLFile)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetXMLFile()", sFuncName)
                DelNodes_XMLFile(sXMLFile, oDICompany, oDITargetComp, String.Empty)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating Item Master " & sItemcode & " in Target DB : " & oDITargetComp.CompanyDB.ToString, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetBusinessObjectFromXML", sFuncName)

                oItemsTarget = oDITargetComp.GetBusinessObjectFromXML(sXMLFile, 0)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling oItemsTarget.Add()", sFuncName)
                lRetCode = oItemsTarget.Add()
                If lRetCode <> 0 Then
                    oDITargetComp.GetLastError(iTemp, sErrDesc)
                    Throw New ArgumentException(iTemp & " : " & sErrDesc)
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successflly Created Item : " & sItemcode & " in Target DB : " & oDITargetComp.CompanyDB.ToString, sFuncName)
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ItemMaster_ADD = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            If System.IO.File.Exists(sXMLFile) Then
                System.IO.File.Delete(sXMLFile)
            End If
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ItemMaster_ADD = RTN_ERROR
        Finally
            If System.IO.File.Exists(sXMLFile) Then
                System.IO.File.Delete(sXMLFile)
            End If
        End Try

    End Function

    Public Function ItemMaster_UPDATE(ByVal oDICompany As SAPbobsCOM.Company, ByVal oDITargetComp As SAPbobsCOM.Company, _
                                        ByVal sItemcode As String, ByVal sTargetItemCode As String, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   ItemMaster_UPDATE()
        '   Purpose     :   This function will Update Item Master in Target Database
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
        Dim oItems As SAPbobsCOM.Items
        Dim oItemsTarget As SAPbobsCOM.Items
        Dim sXMLFile As String = String.Empty
        Dim lRetCode As Long
        Dim iTemp As Integer
        Dim sSQL As String = String.Empty
        Dim oRS As SAPbobsCOM.Recordset
        Dim sTrgXMLFile As String = String.Empty

        Try
            sFuncName = "ItemMaster_UPDATE()"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oItems = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
            oItemsTarget = oDITargetComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

            oRS = oDITargetComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sXMLFile = System.Windows.Forms.Application.StartupPath & "\Item_" & Now.Date.ToString("yyyyMMdd") & Now.ToString("HHmmss") & ".xml"

            sTrgXMLFile = System.IO.Directory.GetCurrentDirectory & "\Trg_DB_Item_" & Now.Date.ToString("yyyyMMdd") & Now.ToString("HHmmss") & ".xml"


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Check Item Exists in Master DB", sFuncName)

            If oItems.GetByKey(sItemcode) Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Saving Item Master as xml file.", sFuncName)

                oDICompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                oItems.SaveXML(sXMLFile)

                oDITargetComp.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                If oItemsTarget.GetByKey(sTargetItemCode) Then
                    oItemsTarget.SaveXML(sTrgXMLFile)
                End If

                DelNodes_XMLFile(sXMLFile, oDICompany, oDITargetComp, sTrgXMLFile)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Check Item Exists in Target DB: " & oDITargetComp.CompanyDB.ToString, sFuncName)
                If oItemsTarget.GetByKey(sTargetItemCode) Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating Item Master " & sItemcode & " in Target DB : " & oDITargetComp.CompanyDB.ToString, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetBusinessObjectFromXML", sFuncName)
                    oItemsTarget = oDITargetComp.GetBusinessObjectFromXML(sXMLFile, 0)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling oItemsTarget.Update()", sFuncName)
                    lRetCode = oItemsTarget.Update()
                    If lRetCode <> 0 Then
                        oDITargetComp.GetLastError(iTemp, sErrDesc)
                        Throw New ArgumentException(iTemp & " : " & sErrDesc)
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Successflly Updated Item : " & sItemcode & " in Target DB : " & oDITargetComp.CompanyDB.ToString, sFuncName)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Item : " & sItemcode & " doesn't exists in Target DB : " & oDITargetComp.CompanyDB.ToString, sFuncName)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ItemMaster_UPDATE = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            If System.IO.File.Exists(sXMLFile) Then
                System.IO.File.Delete(sXMLFile)
            End If

            If System.IO.File.Exists(sTrgXMLFile) Then
                System.IO.File.Delete(sTrgXMLFile)
            End If

            ItemMaster_UPDATE = RTN_ERROR
        Finally
            If System.IO.File.Exists(sXMLFile) Then
                System.IO.File.Delete(sXMLFile)
            End If

            If System.IO.File.Exists(sTrgXMLFile) Then
                System.IO.File.Delete(sTrgXMLFile)
            End If

        End Try
    End Function

    Public Function SyncItemMaster(ByVal sItemCode As String, _
                                   ByVal oDICompany As SAPbobsCOM.Company, _
                                   ByVal oDITargetComp As SAPbobsCOM.Company, _
                                   ByRef sErrDesc As String)

        ' **********************************************************************************
        '   Function    :   SyncItemMaster()
        '   Purpose     :   This function will Sycn Item Master Data 
        '               
        '   Parameters  :   Byval sItemCode as String
        '                       sItemCode= Item Code
        '                   Byval oDICompany As SAPbobsCOM.Company
        '                      oDICompany =  set the SAP DI Company Object
        '                   Byval oDITargetComp As SAPbobsCOM.Company
        '                      oDITargetComp =  set the SAP DI Company Object
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   SRI
        '   Date        :   3 May 2013
        '   Change      :   
        ' ***********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim oDs As New DataSet
        Dim sSQL As String = String.Empty
        Dim oRs As SAPbobsCOM.Recordset
        Dim bIsError As Boolean = False
        Dim sTargetItemCode As String = String.Empty
        Dim oRsTarget As SAPbobsCOM.Recordset
        Dim sSyncItemCode As String = String.Empty

        Try
            sFuncName = "SyncItemMaster()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            oRs = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsTarget = oDITargetComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


            sSQL = "SELECT U_SYNCCODE FROM OITM WHERE ITEMCODE='" & sItemCode & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSQL, sFuncName)
            oRs.DoQuery(sSQL)

            sSyncItemCode = oRs.Fields.Item(0).Value.ToString

            sSQL = "SELECT ITEMCODE FROM OITM WHERE U_SYNCCODE='" & sSyncItemCode & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSQL, sFuncName)
            oRsTarget.DoQuery(sSQL)

            sTargetItemCode = oRsTarget.Fields.Item(0).Value.ToString

            If oRsTarget.EoF Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ItemMaster_ADD", sFuncName)
                If ItemMaster_ADD(oDICompany, oDITargetComp, sItemCode, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ItemMaster_UPDATE", sFuncName)
                If ItemMaster_UPDATE(oDICompany, oDITargetComp, sItemCode, sTargetItemCode, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            SyncItemMaster = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            SyncItemMaster = RTN_ERROR
        Finally

        End Try
    End Function

    Private Sub DelNodes_XMLFile(ByRef sXMLFile As String, _
                                ByVal oDICompany As SAPbobsCOM.Company, ByVal oDITargetComp As SAPbobsCOM.Company, ByVal sTrgXMLfile As String)

        Dim oXMLDoc As New XmlDocument()
        Dim oNodelist As XmlNodeList
        Dim oNode As XmlNode
        Dim sItemcode As String = String.Empty
        Dim sFuncName As String = String.Empty
        Dim oRs As SAPbobsCOM.Recordset
        Dim sSQL As String = String.Empty
        Dim iPriceListNum As Integer
        Dim oTrgXMLDoc As New XmlDocument()
        Dim sXML As String = String.Empty
        Dim sTrgXML As String = String.Empty

        Try
            sFuncName = "DelNodes_XMLFile"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function..", sFuncName)

            oRs = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oXMLDoc.Load(sXMLFile)

            If Not sTrgXMLfile = String.Empty Then
                oTrgXMLDoc.Load(sTrgXMLfile)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Delete OITW Nodes in the XML Files..", sFuncName)
                oNodelist = oXMLDoc.SelectNodes("/BOM/BO/OITW")
                ' Delete OITW Nodes in the XML Files
                For Each oNode In oNodelist
                    oNode.ParentNode.RemoveChild(oNode)
                    oXMLDoc.Save(sXMLFile)
                Next
                'Save XML as String
                sXML = oXMLDoc.OuterXml
                'Get OITW NodeList
                oNodelist = oTrgXMLDoc.SelectNodes("/BOM/BO/OITW")
                'Save OITW NodeList as String
                sTrgXML = oNodelist.Item(0).OuterXml
                'Concatenate TargetDB OITW warehouse details to orignal XML Doc and Save
                sXML = sXML.Replace("</BO></BOM>", String.Empty)
                sXML = sXML & sTrgXML & "</BO></BOM>"
                oXMLDoc.LoadXml(sXML)
            Else
                'oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OITM/row/ItemCode")
                'If Not IsNothing(oNode) Then
                '    oNode.ParentNode.RemoveChild(oNode)
                '    oXMLDoc.Save(sXMLFile)
                'End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Delete OITW Nodes in the XML Files..", sFuncName)
                oNodelist = oXMLDoc.SelectNodes("/BOM/BO/OITW")
                ' Delete OITW Nodes in the XML Files
                For Each oNode In oNodelist
                    oNode.ParentNode.RemoveChild(oNode)
                    oXMLDoc.Save(sXMLFile)
                Next
            End If

            oXMLDoc.Save(sXMLFile)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating Usersing in XML file..", sFuncName)
            ' Change Usersign
            oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OITM/row/UserSign")
            oNode.InnerText = CInt(oDICompany.UserSignature)
            oXMLDoc.Save(sXMLFile)

            ' Delete StockValue/Balances

            oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OITM/row/OpenBlnc")
            If Not IsNothing(oNode) Then
                oNode.ParentNode.RemoveChild(oNode)
                oXMLDoc.Save(sXMLFile)
            End If

            oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OITM/row/OnHand")
            If Not IsNothing(oNode) Then
                oNode.ParentNode.RemoveChild(oNode)
                oXMLDoc.Save(sXMLFile)
            End If

            oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OITM/row/IsCommited")
            If Not IsNothing(oNode) Then
                oNode.ParentNode.RemoveChild(oNode)
                oXMLDoc.Save(sXMLFile)
            End If

            oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OITM/row/OnOrder")
            If Not IsNothing(oNode) Then
                oNode.ParentNode.RemoveChild(oNode)
                oXMLDoc.Save(sXMLFile)
            End If

            oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OITM/row/LastPurPrc")
            If Not IsNothing(oNode) Then
                oNode.ParentNode.RemoveChild(oNode)
                oXMLDoc.Save(sXMLFile)
            End If

            oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OITM/row/LstEvlPric")
            If Not IsNothing(oNode) Then
                oNode.ParentNode.RemoveChild(oNode)
                oXMLDoc.Save(sXMLFile)
            End If

            oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OITM/row/ExitPrice")
            If Not IsNothing(oNode) Then
                oNode.ParentNode.RemoveChild(oNode)
                oXMLDoc.Save(sXMLFile)
            End If


            oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OITM/row/StockValue")
            If Not IsNothing(oNode) Then
                oNode.ParentNode.RemoveChild(oNode)
                oXMLDoc.Save(sXMLFile)
            End If

            oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OITM/row/AvgPrice")
            If Not IsNothing(oNode) Then
                oNode.ParentNode.RemoveChild(oNode)
                oXMLDoc.Save(sXMLFile)
            End If

            oNodelist = oXMLDoc.SelectNodes("/BOM/BO/OITW/row/AvgPrice")
            For Each oNode In oNodelist
                If Not IsNothing(oNode) Then
                    oNode.ParentNode.RemoveChild(oNode)
                    oXMLDoc.Save(sXMLFile)
                End If
            Next

            oNode = oXMLDoc.SelectSingleNode("/BOM/BO/OITM/row/PricingPrc")
            If Not IsNothing(oNode) Then
                oNode.ParentNode.RemoveChild(oNode)
                oXMLDoc.Save(sXMLFile)
            End If


            '1. GET ItemGroup code from TargetDB based on Master DB ItemGroupName and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/ItmsGrpCod", "OITB", "ItmsGrpCod", "ItmsGrpNam", True, oXMLDoc, sXMLFile)

            '2. Get PriceListNum from TargetDB base on Master DB PriceList Name and Update
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating PriceListNum in XML file..", sFuncName)
            oNodelist = oXMLDoc.SelectNodes("/BOM/BO/ITM1/row/PriceList")

            For Each oNode In oNodelist
                iPriceListNum = CInt(oNode.InnerText)
                sSQL = "SELECT ListNum from  [" & oDITargetComp.CompanyDB.ToString & "].[dbo].OPLN WHERE ListName in (select ListName from OPLN WHERE ListNum=" & iPriceListNum & ")"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL Query" & sSQL, sFuncName)
                oRs.DoQuery(sSQL)
                If Not oRs.EoF Then oNode.InnerText = oRs.Fields.Item("ListNum").Value
                oXMLDoc.Save(sXMLFile)
            Next

            '3. GET ShipType code from TargetDB based on Master DB ShipTypeName and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/ShipType", "OSHP", "TrnspCode", "TrnspName", True, oXMLDoc, sXMLFile)

            '4. GET Manufacturer code from TargetDB based on Master DB ManufacturerName and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/FirmCode", "OMRC", "FirmCode", "FirmName", True, oXMLDoc, sXMLFile)

            '5. GET Custom Group code from TargetDB based on Master DB Custom Group Name and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/CstGrpCode", "OARG", "CstGrpCode", "CstGrpName", True, oXMLDoc, sXMLFile)

            '6. GET OrderInterval from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/OrdrIntrvl", "OCYC", "Code", "Name", True, oXMLDoc, sXMLFile)

            '7. GET Purchasing Heigth Unit 1  from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/BHght1Unit", "OLGT", "UnitCode", "UnitDisply", True, oXMLDoc, sXMLFile)

            '8. GET Purchasing Heigth Unit 2  from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/BHght2Unit", "OLGT", "UnitCode", "UnitDisply", True, oXMLDoc, sXMLFile)

            '9. GET Purchasing Length Unit 1 from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/BLen1Unit", "OLGT", "UnitCode", "UnitDisply", True, oXMLDoc, sXMLFile)

            '10. GET Purchasing Length Unit 2 from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/BLen2Unit", "OLGT", "UnitCode", "UnitDisply", True, oXMLDoc, sXMLFile)

            '11. GET Purchasing Width Unit 1  from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/BWdth1Unit", "OLGT", "UnitCode", "UnitDisply", True, oXMLDoc, sXMLFile)

            '12. GET Purchasing Width Unit 2  from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/BWdth2Unit", "OLGT", "UnitCode", "UnitDisply", True, oXMLDoc, sXMLFile)

            '13. GET Purchasing Weight Unit 1  from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/BWght1Unit", "OLGT", "UnitCode", "UnitDisply", True, oXMLDoc, sXMLFile)

            '14. GET Purchasing Weight Unit 2  from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/BWght2Unit", "OLGT", "UnitCode", "UnitDisply", True, oXMLDoc, sXMLFile)

            '15. GET Purchasing Volume Unit from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/BVolUnit", "OLGT", "UnitCode", "VolDisply", True, oXMLDoc, sXMLFile)

            '------- Sales L/W/H/Volume/Weight

            '16. GET Purchasing Heigth Unit 1  from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/SHght1Unit", "OLGT", "UnitCode", "UnitDisply", True, oXMLDoc, sXMLFile)

            '17. GET Purchasing Heigth Unit 2  from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/SHght2Unit", "OLGT", "UnitCode", "UnitDisply", True, oXMLDoc, sXMLFile)

            '18. GET Purchasing Length Unit 1 from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/SLen1Unit", "OLGT", "UnitCode", "UnitDisply", True, oXMLDoc, sXMLFile)

            '19. GET Purchasing Length Unit 2 from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/SLen2Unit", "OLGT", "UnitCode", "UnitDisply", True, oXMLDoc, sXMLFile)

            '20. GET Purchasing Width Unit 1  from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/SWdth1Unit", "OLGT", "UnitCode", "UnitDisply", True, oXMLDoc, sXMLFile)

            '21. GET Purchasing Width Unit 2  from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/SWdth2Unit", "OLGT", "UnitCode", "UnitDisply", True, oXMLDoc, sXMLFile)

            '22. GET Purchasing Weight Unit 1  from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/SWght1Unit", "OLGT", "UnitCode", "UnitDisply", True, oXMLDoc, sXMLFile)

            '23. GET Purchasing Weight Unit 2  from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/SWght2Unit", "OLGT", "UnitCode", "UnitDisply", True, oXMLDoc, sXMLFile)

            '24. GET Purchasing Volume Unit from TargetDB based on Master DB and Update 
            UpdateXML(oDICompany, oDITargetComp, "/BOM/BO/OITM/row/SVolUnit", "OLGT", "UnitCode", "VolDisply", True, oXMLDoc, sXMLFile)

            oNode = Nothing
            oXMLDoc = Nothing

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch ex As Exception

        End Try

    End Sub

#End Region

End Module
