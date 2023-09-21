dim uftObject

Set uftObject = CreateObject("QuickTest.Application")

uftObject.Visible = True

MsgBox "UFT Start"

uftObject.Launch

wscript.sleep 4000

uftObject.Open "C:\Users\gauri\Documents\UFT One\Group1_UFTAssignment"

Set result = CreateObject("QuickTest.RunResultsOptions")
result.Resultslocation = "C:\Users\gauri\Documents\UFT One"

uftObject.Test.Run result 
MsgBox "UFT Finished"

uftObject.Test.LastRunResults

uftObject.Test.Save
