Run the Pester test suite and analyze the results.

## Steps

1. Run Pester with detailed output:
```
!pwsh -Command "Invoke-Pester -Output Detailed -PassThru | ConvertTo-Json -Depth 5"
```

2. If all tests pass, confirm the count and stop.

3. If any tests fail, for each failure:
   - State the test name and file
   - Show the expected vs. actual values
   - Read the relevant source function to understand the failure
   - Explain why the test is failing
   - Suggest a fix — in the source code if it's a bug, or in the test if the test is wrong

4. If Pester is not installed or the command fails:
   - Check whether the Pester module is available: `Get-Module -ListAvailable Pester`
   - If missing, tell me and suggest `Install-Module Pester -Force -Scope CurrentUser`
   - If it's a different error, show the full error output

$ARGUMENTS
