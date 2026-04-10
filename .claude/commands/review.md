Review the current branch against main. Focus on correctness, not style.

## Steps

1. Run the diff:
```
!git diff main...HEAD --stat
```

2. Get the full diff for changed files:
```
!git diff main...HEAD
```

3. Review each changed file for:
   - **Bugs** — logic errors, off-by-one, null/empty checks, unhandled error paths
   - **PowerShell anti-patterns** — using `Write-Host` where `Write-Output` or `Write-Verbose` belongs, missing `[CmdletBinding()]`, untyped parameters, swallowed exceptions
   - **Pipeline contract violations** — functions that should accept pipeline input but don't, or that emit unexpected types
   - **Missing `-WhatIf`/`-Confirm`** on anything that writes, renames, deletes, or transfers
   - **Security** — hardcoded paths to seedbox/credentials, secrets in source

4. Summarize findings as a numbered list. For each issue:
   - File and function name
   - What the problem is
   - Suggested fix (with code if straightforward)

5. If nothing notable is found, say so — don't invent issues.

$ARGUMENTS
