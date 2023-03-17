Set oShell = CreateObject("WScript.Shell")

'Set WshNetwork = wscript.Arguments
'MsgBox()
'For Each arg In args
'  MsgBox(arg)
'Next
'Set obj = oShell.Exec("cmd /c" & "echo aaaa")
Set obj = oShell.Exec("cmd /c" & "curl -H ""Authorization: Basic Y2hvcm5kOmVtcGlyaWMwMSoxMTA0MTE="" -H ""Content-Type: application/json"" -X POST https://jira-test.nexinsure.org/rest/atm/1.0/testrun/" & TestArgs("TestCycleKey") & "/testcase/TESTMGMT-T6/testresult -d ""{""""status"""": """"Pass"""",""""comment"""": """"The test has failed on some automation tool procedure."""",""""actualStartDate"""": """"" & TestArgs("ActualStartDate") & """"",""""actualEndDate"""": """"" & TestArgs("ActualStartDate") & """""}""")
'Do While Not obj.StdOut.AtEndOfStream
sText = obj.StdOut.ReadLine()
MsgBox(sText)
Set oShell = Nothing @@ script infofile_;_ZIP::ssf3.xml_;_


