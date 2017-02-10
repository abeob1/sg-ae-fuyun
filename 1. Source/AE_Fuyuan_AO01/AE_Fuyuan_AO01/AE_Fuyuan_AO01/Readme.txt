For Every Activity, must put this code of line.

If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Description", sFuncName)

************************************************************************************************
									DISPLAY WAITING STATUS
************************************************************************************************

If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DisplayStatus()", sFuncName)
Call DisplayStatus(oForm, "Please wait....", sErrDesc)

If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EndStatus()", sFuncName)
Call EndStatus(sErrDesc)

************************************************************************************************
									FUNCTION TEMPLATE
************************************************************************************************

Public Function MyFunction(ByVal oForm As SAPbouiCOM.Form, Byref sErrDesc As String) As Long
  
        'Function   :   MyFunction()
        'Purpose    :   
        'Parameters :   ByVal oForm As SAPbouiCOM.Form
        '                   oForm=Form Type
        '               ByRef sErrDesc As String
        '                   sErrDesc=Error Description to be returned to calling function
        '               
        '                   =
        'Return     :   0 - FAILURE
        '               1 - SUCCESS
        'Author     :   SRI
        'Date       :   30/12/2007
        'Change     :
  
  Dim sFuncName As String
  
    Try
		sFuncName = "MyFunction()"
		If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
    
    
        'Start Function here
        Dim oRS As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
		Dim sSql As String
		ssql=""
		oRS.DoQuery(sSql)

        MyFunction = RTN_SUCCESS
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
    
    Catch exc As Exception
        MyFunction = RTN_ERROR
        sErrDesc = exc.Message
        Call WriteToLogFile(sErrDesc, sFuncName)
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
    Finally
    End Try
End Function
    
************************************************************************************************
							CALLING FUNCTION WITH THROW ARGUMENT EXCEPTION
************************************************************************************************     
If MyFunction(oForm, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
If sBPCode = "" Then Throw New ArgumentException("BP Code not provided.")

************************************************************************************************    
									START/COMMIT/ROLLBACK TRANSACTION
************************************************************************************************    
if StartTransaction_ADO(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
if RollBackTransaction_ADO(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
if CommitTransaction_ADO(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

if StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
if RollBackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
if CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc


************************************************************************************************
								EVENT LIST
************************************************************************************************ 
Application events - events triggered by actions on the Company database of the SAP Business One application. 
MenuEvent events - events triggered by actions on the main menu and the menu bar.
ItemEvent events - events triggered by actions on forms and items.
ProgressBar events - events occurring while the progress bar is activated.
StatusBar event - event triggered by sending a message to the application status bar.
Print event -	End-user clicks on Print or Print Preview icons
				End-user sends a document to print using the Document Printing option
				The application Document Generation Wizard sends a document to print
				You can use this event as an "exit" point for integrating third-party reporting tools instead of the SAP Business One Document Editor.
Form Data events - 
				Form Data events occurs when the application performs the following actions on forms connected to business objects:
				Add
				Update
				Delete
				Load form data via browse, link button, or find
Right-click event - occurs when an end-user presses the right mouse button on a specific item in the application's forms.

*************************************************************************************************
									SAP FORM MODE
************************************************************************************************

If pVal.ItemUID = "1" And pVal.Before_Action = False Then
    Select Case pVal.FormMode
        Case 0 'search
            p_oSBOApplication.MessageBox("Search Mode")
        Case 1 'ok
            p_oSBOApplication.MessageBox("OK Mode")
        Case 2 'update
            p_oSBOApplication.MessageBox("Update Mode")
        Case 3
            p_oSBOApplication.MessageBox("Add New Mode")
    End Select
End If
                    
                    
************************************************************************************************                   
									GET UDF FIELD
************************************************************************************************
Dim oFormLK As SAPbouiCOM.Form
oFormLK = p_oSBOApplication.Forms.GetForm("-" & oForm.TypeEx, oForm.TypeCount)
Dim sPlantID As String = oFormLK.Items.Item("U_AB_PlaCd").Specific.selected.value


************************************************************************************************
									MATRIX LOOPING
************************************************************************************************

For iCnt = 1 To oForm.Items.Item(sMatrixName).Specific.visualRowCount
    sVariable = oForm.Items.Item(sMatrixName).Specific.columns.item("AB_Var").cells.item(iCnt).specific.value
    dValue = oForm.Items.Item(sMatrixName).Specific.columns.item("AB_Value").cells.item(iCnt).specific.value
Next

************************************************************************************************
									GET COMPANY OBJECT HERE...
************************************************************************************************
p_oSBOApplication.Company.


- Check for DCn Value, make sure user got enter before save/calculate
- user cannot click the DCn value on matrix.


