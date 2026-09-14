<#
.SYNOPSIS
    Executes the runner's endpoint observation - Invoke-FaObservedEndpoint and the
    Invoke-FaEndpoint wrapper - lifted out of the runner by AST, against a stand-in
    Excel whose Run models the production automation seam: PCCM_AutomationBegin
    clears the recorded prompt and result, the endpoint records both, the two
    readers return them. Excel is never started. IT ASSERTS NOTHING: it prints
    FLOW|<case>|<text> lines and the Python control decides.
#>
param([string]$Runner)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$source = Get-Content -LiteralPath $Runner -Raw
$ast = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$null, [ref]$null)
foreach ($name in @('Invoke-FaObservedEndpoint', 'Invoke-FaEndpoint')) {
    $definition = $ast.FindAll({ param($node)
        ($node -is [System.Management.Automation.Language.FunctionDefinitionAst]) -and ($node.Name -eq $name) }, $true)
    if (@($definition).Count -ne 1) { throw ('the runner defines ' + $name + ' ' + [string]@($definition).Count + ' times') }
    . ([scriptblock]::Create(@($definition)[0].Extent.Text))
}

# THE STAND-IN SEAM, modelled on modAppState: Begin clears everything, the
# endpoint records a prompt and a result, the readers hand them back.
$script:Seam = @{ Active = $false; Reply = $false; Failpoint = ''; Prompt = ''; Result = ''; Begins = 0 }
$stub = New-Object PSObject
$stub | Add-Member -MemberType ScriptMethod -Name Run -Value {
    param($Procedure, $Arg1, $Arg2)
    switch ($Procedure) {
        'PCCM_AutomationBegin' {
            $script:Seam.Active = $true; $script:Seam.Reply = [bool]$Arg1; $script:Seam.Failpoint = [string]$Arg2
            $script:Seam.Prompt = ''; $script:Seam.Result = ''; $script:Seam.Begins = $script:Seam.Begins + 1
            return $null
        }
        'PCCM_ResetResults' {
            $script:Seam.Prompt = 'Reset Results clears every published result and keeps every input.'
            if ($script:Seam.Reply) { $script:Seam.Result = 'OK|Results reset.' } else { $script:Seam.Result = 'OK|' }
            return $null
        }
        'PCCM_AutomationResult' { return $script:Seam.Result }
        'PCCM_AutomationPrompt' { return $script:Seam.Prompt }
        default { throw ('the stand-in seam does not model ' + [string]$Procedure) }
    }
}

function Emit { param([string]$Case, [string]$Text) Write-Output ('FLOW|' + $Case + '|' + $Text) }

$observed = Invoke-FaObservedEndpoint -Excel $stub -Operation 'PCCM_ResetResults' -ConfirmReply $false -FailAfterStage 'SomeStage'
Emit 'observed.declined' ('result=' + $observed.Result + '|prompt=' + $observed.Prompt)
Emit 'after.cleanup' ('seam.prompt=<' + $script:Seam.Prompt + '>|seam.result=<' + $script:Seam.Result + '>|reply=' + [string]$script:Seam.Reply + '|failpoint=<' + $script:Seam.Failpoint + '>|begins=' + [string]$script:Seam.Begins)
Emit 'post-cleanup.reader' ('prompt=<' + [string]$stub.Run('PCCM_AutomationPrompt') + '>')
$wrapped = Invoke-FaEndpoint -Excel $stub -Operation 'PCCM_ResetResults'
Emit 'wrapper.confirmed' ('result=' + $wrapped + '|type=' + $wrapped.GetType().Name)
