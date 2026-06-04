import re

with open("Matcher.bas", "r") as f:
    content = f.read()

content = re.sub(
    r"On Error Resume Next: cell\.Validation\.Delete: On Error GoTo 0\s*\n\s*With cell\.Validation\s*\n\s*\.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=listString\s*\n\s*\.IgnoreBlank = True\s*\n\s*\.InCellDropdown = True\s*\n\s*\.ShowError = False\s*\n\s*End With",
    "ApplyDropdownToCell cell, listString",
    content
)

content = re.sub(
    r"On Error Resume Next\s*\n\s*ws\.Cells\(rowNum, \"B\"\)\.Validation\.Delete\s*\n\s*On Error GoTo 0\s*\n\s*With ws\.Cells\(rowNum, \"B\"\)\.Validation\s*\n\s*\.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=matchStr & \",Reject AI Guess\"\s*\n\s*\.IgnoreBlank = True\s*\n\s*\.InCellDropdown = True\s*\n\s*\.ShowError = False\s*\n\s*End With",
    "ApplyDropdownToCell ws.Cells(rowNum, \"B\"), matchStr & \",Reject AI Guess\"",
    content
)

content = re.sub(
    r"On Error Resume Next\s*\n\s*ws\.Cells\(rRow, \"B\"\)\.Validation\.Delete\s*\n\s*On Error GoTo 0\s*\n\s*With ws\.Cells\(rRow, \"B\"\)\.Validation\s*\n\s*\.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=rMatch & \",Reject AI Guess\"\s*\n\s*\.IgnoreBlank = True\s*\n\s*\.InCellDropdown = True\s*\n\s*\.ShowError = True\s*\n\s*End With",
    "ApplyDropdownToCell ws.Cells(rRow, \"B\"), rMatch & \",Reject AI Guess\"",
    content
)

with open("Matcher.bas", "w") as f:
    f.write(content)
