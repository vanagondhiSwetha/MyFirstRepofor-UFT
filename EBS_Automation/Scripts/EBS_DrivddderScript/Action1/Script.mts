'***************************** 
'Title:EBS DriverScript
'Created by-Date(MM/DD/YY): Swetha Vanagondhi-07/02/2019
'Modified by-Date(MM/DD/YY):
'Comments:
'Description: Driver Script used to run all EBS scripts 
'***************************** 
Public TestcaseName
Public UnknownTestcaseName

Set TestSuite_Excel = CreateObject("Excel.Application") 
Set objWorkbook = TestSuite_Excel.Workbooks.Open("C:\EBS_Oracle\EBS_Automation\TestSuite.xlsx")
Set objWorksheet = objWorkbook.Worksheets("EBSTestSuit")
Rowcount = objWorksheet.UsedRange.Rows.count
TestSuite_Excel.Visible = True 

For J = 2 to Rowcount

'j = i+1
  RunStatus = objWorksheet.Cells(j, 1).Value
  TestcaseName = objWorksheet.Cells(j, 2).value
  UnknownTestcaseName = objWorksheet.Cells(j, 7).value
  Environment.Value("ReceivngTestName") = TestcaseName
  Environment.Value("ReceivngUnknownTestName") = UnknownTestcaseName

If ucase(RunStatus)=ucase("yes") then
     strFeatureFile = ""&TestcaseName&".vbs"    
     ReceivingTestCase = Split(TestcaseName,"_")
    Environment.Value("TestCaseRownumber") = J
 

     If ReceivingTestCase(0) = "Receiving-YK" or ReceivingTestCase(0) = "Receiving-PSEUDO-YK" Then       
	   LoadFunctionLibrary("C:\EBS_Oracle\EBS_Automation\EBS-Scripts\Receiving_YK_WS.vbs")	        
     ElseIf ReceivingTestCase(0) = "Receiving-LV" or ReceivingTestCase(0) = "Receiving-PSEUDO-LV" Then
       LoadFunctionLibrary("C:\EBS_Oracle\EBS_Automation\EBS-Scripts\Receiving_LV_All_Good.vbs")  
     Else
       LoadFunctionLibrary("C:\EBS_Oracle\EBS_Automation\EBS-Scripts\"&strFeatureFile&"")      
	 End If      
	

	If intFail>0 Then
    	objWorksheet.Cells(j,3).Value = "FAIL" 
    	'objWorksheet.Cells(j,3).Interior.Color = vbRED
    	objWorksheet.Cells(j,3).Font.Color = vbRED
    	intFail = ""
     Else
       objWorksheet.Cells(j,3).Value = "PASS"
    	'objWorksheet.Cells(j,3).Interior.Color = vbGREEN
    	objWorksheet.Cells(j,3).Font.Color = vbGREEN
    End If

    objWorksheet.Cells(j,4).Value = "Html Result link"
	range1 = "D"&j 
	range2 = "E"&j
	Set objRange1 = TestSuite_Excel.Range(range1) 
	Set objRange2 = TestSuite_Excel.Range(range2) 
	Set objLink = objWorksheet.Hyperlinks.Add(objRange1,strFilePath)
	Set objLink = objWorksheet.Hyperlinks.Add(objRange2,ExcelReport)
	objWorksheet.Cells(j,6).Value = now()

End if 

Next
'CloseFile ()
objWorkbook.Save

Set TestSuite_Excel = Nothing

 @@ hightlight id_;_2954_;_script infofile_;_ZIP::ssf113.xml_;_
 
 
 
 
 @@ hightlight id_;_678_;_script infofile_;_ZIP::ssf121.xml_;_
