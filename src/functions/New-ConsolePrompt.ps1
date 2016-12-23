Function New-ConsolePrompt {
    param (
        [Parameter(Position=0)]
        [string]$UserPrompt
    )
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Yes."
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "No."
    $ContinueBuildPrompt = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
    if (($host.ui.PromptForChoice('', $UserPrompt, $ContinueBuildPrompt, 0)) -eq 0) {
        $true
    }
    else {
        $false
    }
}