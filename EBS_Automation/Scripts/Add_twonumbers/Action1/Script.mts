'Adding required repositories and function libraries
LoadFunctionLibrary "C:\EBS_Oracle\EBS_Automation\Functional_Library\Common_Functions.qfl"
LoadLibraries()



path=InitializeReporter_Web()
OpenFile(path)

Scenario = TestName
Createcomheadervalues()

a=5
b=7
c=a+b

Addguisteps "Adding numbers","validating results","Result"&"="&c,"c"&"="&c,"Pass" 
Reporter.ReportEvent micPass, "Adding numbers", "Result"&"="&c
