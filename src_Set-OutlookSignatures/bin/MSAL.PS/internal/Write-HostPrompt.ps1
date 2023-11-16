<#
.SYNOPSIS
    Displays a PowerShell prompt for multiple fields or multiple choices.
.DESCRIPTION
    Displays a PowerShell prompt for multiple fields or multiple choices.
.EXAMPLE
    PS C:\>Write-HostPrompt "Prompt Caption" -Fields "Field 1", "Field 2"
    Display simple prompt for 2 fields.
.EXAMPLE
    PS C:\>$IntegerField = New-Object System.Management.Automation.Host.FieldDescription -ArgumentList "Integer Field" -Property @{ HelpMessage = "Help Message for Integer Field" }
    PS C:\>$IntegerField.SetParameterType([int[]])
    PS C:\>$DateTimeField = New-Object System.Management.Automation.Host.FieldDescription -ArgumentList "DateTime Field" -Property @{ HelpMessage = "Help Message for DateTime Field" }
    PS C:\>$DateTimeField.SetParameterType([datetime])
    PS C:\>Write-HostPrompt "Prompt Caption" "Prompt Message" -Fields $IntegerField, $DateTimeField
    Display prompt for 2 type-specific fields, with int field being an array.
.EXAMPLE
    PS C:\>Write-HostPrompt "Prompt Caption" -Choices "Choice &1", "Choice &2"
    Display simple prompt with 2 choices.
.EXAMPLE
    PS C:\>Write-HostPrompt "Prompt Caption" "Prompt Message" -DefaultChoice 2 -Choices @(
        New-Object System.Management.Automation.Host.ChoiceDescription -ArgumentList "&1`bChoice one" -Property @{ HelpMessage = "Help Message for Choice 1" }
        New-Object System.Management.Automation.Host.ChoiceDescription -ArgumentList "&2`bChoice two" -Property @{ HelpMessage = "Help Message for Choice 2" }
    )
    Display prompt with 2 choices and help messages that defaults to the second choice.
.EXAMPLE
    PS C:\>Write-HostPrompt "Prompt Caption" "Choose a number" -Choices "Menu Item A", "Menu Item B", "Menu Item C" -HelpMessages "Menu Item A Needs Help", "Menu Item B Needs More Help",, "Menu Item C Needs Crazy Help" -NumberedHotKeys
    Display prompt with 3 choices and help message that are automatically numbered.
.INPUTS
    System.Management.Automation.Host.FieldDescription
    System.Management.Automation.Host.ChoiceDescription
.OUTPUTS
    System.Collections.Generic.Dictionary[System.String,System.Management.Automation.PSObject]
    System.Int32
.LINK
    https://github.com/jasoth/Utility.PS
#>
function Write-HostPrompt {
    [CmdletBinding()]
    param
    (
        # Caption to preceed or title the prompt.
        [Parameter(Mandatory = $true, Position = 1)]
        [string] $Caption,
        # A message that describes the prompt.
        [Parameter(Mandatory = $false, Position = 2)]
        [string] $Message,
        # The fields in the prompt.
        [Parameter(Mandatory = $true, ParameterSetName = 'Fields', Position = 3, ValueFromPipeline = $true)]
        [System.Management.Automation.Host.FieldDescription[]] $Fields,
        # The choices the shown in the prompt.
        [Parameter(Mandatory = $true, ParameterSetName = 'Choices', Position = 3, ValueFromPipeline = $true)]
        [System.Management.Automation.Host.ChoiceDescription[]] $Choices,
        # Specifies a help message for each field or choice.
        [Parameter(Mandatory = $false, Position = 4)]
        [string[]] $HelpMessages = @(),
        # The index of the label in the choices to make default.
        [Parameter(Mandatory = $false, ParameterSetName = 'Choices', Position = 5)]
        [int] $DefaultChoice,
        # Use numbered hot keys (aka "keyboard accelerator") for each choice.
        [Parameter(Mandatory = $false, ParameterSetName = 'Choices', Position = 6)]
        [switch] $NumberedHotKeys
    )

    begin {
        ## Create list to capture multiple fields or multiple choices.
        [System.Collections.Generic.List[System.Management.Automation.Host.FieldDescription]] $listFields = New-Object System.Collections.Generic.List[System.Management.Automation.Host.FieldDescription]
        [System.Collections.Generic.List[System.Management.Automation.Host.ChoiceDescription]] $listChoices = New-Object System.Collections.Generic.List[System.Management.Automation.Host.ChoiceDescription]
    }

    process {
        switch ($PSCmdlet.ParameterSetName) {
            'Fields' {
                for ($iField = 0; $iField -lt $Fields.Count; $iField++) {
                    if ($iField -lt $HelpMessages.Count -and $HelpMessages[$iField]) { $Fields[$iField].HelpMessage = $HelpMessages[$iField] }
                    $listFields.Add($Fields[$iField])
                }
            }
            'Choices' {
                for ($iChoice = 0; $iChoice -lt $Choices.Count; $iChoice++) {
                    if ($NumberedHotKeys) { $Choices[$iChoice] = New-Object System.Management.Automation.Host.ChoiceDescription -ArgumentList "&$($iChoice+1)`b$($Choices[$iChoice].Label)" -Property @{ HelpMessage = $Choices[$iChoice].HelpMessage } }
                    #elseif (!$Choices[$iChoice].Label.Contains('&')) { $Choices[$iChoice] = New-Object System.Management.Automation.Host.ChoiceDescription -ArgumentList "&$($Choices[$iChoice].Label)" -Property @{ HelpMessage = $Choices[$iChoice].HelpMessage } }
                    if ($iChoice -lt $HelpMessages.Count -and $HelpMessages[$iChoice]) { $Choices[$iChoice].HelpMessage = $HelpMessages[$iChoice] }
                    $listChoices.Add($Choices[$iChoice])
                }
            }
        }
    }

    end {
        try {
            switch ($PSCmdlet.ParameterSetName) {
                'Fields' { return $Host.UI.Prompt($Caption, $Message, $listFields.ToArray()) }
                'Choices' { return $Host.UI.PromptForChoice($Caption, $Message, $listChoices.ToArray(), $DefaultChoice - 1) + 1 }
            }
        }
        catch [System.Management.Automation.PSInvalidOperationException] {
            ## Write Non-Terminating Error When In Non-Interactive Mode.
            Write-Error -ErrorRecord $_ -CategoryActivity $MyInvocation.MyCommand
        }
    }
}
