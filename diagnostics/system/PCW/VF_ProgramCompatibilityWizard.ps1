# Copyright © 2008, Microsoft Corporation. All rights reserved.


#if this environment variable is set, we say that we don't detect the problem anymore so it will
#show as fixed in the final screen

if ($Env:AppFixed -eq $true)
{
    Update-DiagRootCause -id "RC_IncompatibleApplication" -Detected $false
}

#RS_ProgramCompatibilityWizard
#rparsons - 05 May 2008
