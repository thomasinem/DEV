  # Copyright © 2008, Microsoft Corporation. All rights reserved.


#This is passed from the troubleshooter via 'Add-DiagRootCause'
PARAM($targetPath, $appName)

#RS_ProgramCompatibilityWizard
#rparsons - 05 May 2008
#rfink    - 01 Sept 2008 - rewrite to support dynamic choices 

#set-psdebug -strict -trace 0

#change HKLM\Software\Windows NT\CurrentVersion\AppCompatFlags\CompatTS EnableTracing(DWORD) to 1
#if you want to enable tracing
$SpewTraceToDesktop = $false

Import-LocalizedData -BindingVariable CompatibilityStrings -FileName CL_LocalizationData

#Compatibility modes
$CompatibilityModes = new-Object System.Collections.Hashtable
$CompatibilityModes.Add("Version_WIN8RTM", "WIN8RTM")
$CompatibilityModes.Add("Version_WIN7RTM", "WIN7RTM")
$CompatibilityModes.Add("Version_WINVISTA2", "VISTASP2")
$CompatibilityModes.Add("Version_WINXP3", "WINXPSP3")
$CompatibilityModes.Add("Version_MSIAUTO", "MSIAUTO")
$CompatibilityModes.Add("Version_UNKNOWN", "WINXPSP3")
$CompatibilityModes.Add("Display_256COLOR", "256COLOR")
$CompatibilityModes.Add("Display_16BITCOLOR", "16BITCOLOR")
$CompatibilityModes.Add("Display_640x480", "640X480")
$CompatibilityModes.Add("Display_HIGHDPIAWARE", "HIGHDPIAWARE")
$CompatibilityModes.Add("Access_RUNASADMIN", "RUNASADMIN")

[string]$RunAsAdminCompatMode = "RUNASADMIN"
[string]$MsiAutoCompatMode = "MSIAUTO"
[string]$AllVersionModes = "WIN8RTM WIN7RTM VISTASP2 WINXPSP3"
[string]$AllDisplayModes = "256COLOR 16BITCOLOR 640X480 HIGHDPIAWARE"
[string]$VistaPlusDisplayMode = "HIGHDPIAWARE"
[string]$AllTSLayers = "{0} {1} {2} {3}" -f $AllVersionModes, $RunAsAdminCompatMode, $MsiAutoCompatMode, $AllDisplayModes

$SupportedModes = new-Object System.Collections.ArrayList
$SupportedModes.AddRange($CompatibilityModes.Values)

#Compatibility mode strings
$CompatibilityModeStrings = new-Object System.Collections.Hashtable
$CompatibilityModeStrings.Add("WIN8RTM", $CompatibilityStrings.Version_Choice_WIN8RTM)
$CompatibilityModeStrings.Add("WIN7RTM", $CompatibilityStrings.Version_Choice_WIN7RTM)
$CompatibilityModeStrings.Add("VISTASP2", $CompatibilityStrings.Version_Choice_WINVISTA2)
$CompatibilityModeStrings.Add("WINXPSP3", $CompatibilityStrings.Version_Choice_WINXPSP3)
$CompatibilityModeStrings.Add("MSIAUTO", $CompatibilityStrings.Version_Choice_MSIAUTO)
$CompatibilityModeStrings.Add("256COLOR", $CompatibilityStrings.Display_Choice_256COLOR)
$CompatibilityModeStrings.Add("16BITCOLOR", $CompatibilityStrings.Display_Choice_16BITCOLOR)
$CompatibilityModeStrings.Add("640X480", $CompatibilityStrings.Display_Choice_640x480)
$CompatibilityModeStrings.Add("HIGHDPIAWARE", $CompatibilityStrings.Display_Choice_HIGHDPIAWARE)

[int]$VersionProblem    = 1
[int]$DisplayProblem    = 2
[int]$RunAsAdminProblem = 4

[int]$problemMask = 0

[string]$spacer = " "
[string]$displaySpacer = ", "
[string]$delimiters = "# "

#Xml constants
[string]$resultSuccess = "Success"
[string]$resultFailure = "Failure"

$problemChoiceXml=@'
<Choices>
    <Choice>
        <Name>ProblemN_Choice_VERSION</Name>
        <Description>ProblemD_Choice_VERSION</Description>
        <Value>VersionProblem</Value>
        <ExtensionPoint />
    </Choice>
    <Choice>
        <Name>ProblemN_Choice_DISPLAY</Name>
        <Description>ProblemD_Choice_DISPLAY</Description>
        <Value>DisplayProblem</Value>
        <ExtensionPoint />
    </Choice>
    <Choice>
        <Name>ProblemN_Choice_ACCESS</Name>
        <Description>ProblemD_Choice_ACCESS</Description>
        <Value>RunAsAdminProblem</Value>
        <ExtensionPoint />
    </Choice>
    <Choice>
        <Name>ProblemN_Choice_UNKNOWN</Name>
        <Description>ProblemD_Choice_UNKNOWN</Description>
        <Value>UnknownProblem</Value>
        <ExtensionPoint />
    </Choice>
</Choices>
'@

$problemChoiceXmlMsi=@'
<Choices>
    <Choice>
        <Name>ProblemN_Choice_VERSION</Name>
        <Description>ProblemD_Choice_VERSION</Description>
        <Value>VersionProblem</Value>
        <ExtensionPoint />
    </Choice>
    <Choice>
        <Name>ProblemN_Choice_DISPLAY</Name>
        <Description>ProblemD_Choice_DISPLAY</Description>
        <Value>DisplayProblem</Value>
        <ExtensionPoint />
    </Choice>
    <Choice>
        <Name>ProblemN_Choice_UNKNOWN</Name>
        <Description>ProblemD_Choice_UNKNOWN</Description>
        <Value>UnknownProblem</Value>
        <ExtensionPoint />
    </Choice>
</Choices>
'@

$versionChoiceXml=@'
<Choices>
    <Choice>
        <Name>Version_Choice_WIN8RTM</Name>
        <Description>VersionD_Choice_ALL</Description>
        <Value>Version_WIN8RTM</Value>
        <ExtensionPoint />
    </Choice>
    <Choice>
        <Name>Version_Choice_WIN7RTM</Name>
        <Description>VersionD_Choice_ALL</Description>
        <Value>Version_WIN7RTM</Value>
        <ExtensionPoint>
            <Default />
        </ExtensionPoint>
    </Choice>
    <Choice>
        <Name>Version_Choice_WINVISTA2</Name>
        <Description>VersionD_Choice_ALL</Description>        
        <Value>Version_WINVISTA2</Value>
        <ExtensionPoint />
    </Choice>
    <Choice>
        <Name>Version_Choice_WINXPSP3</Name>
        <Description>VersionD_Choice_ALL</Description>        
        <Value>Version_WINXP3</Value>
        <ExtensionPoint />
    </Choice>
    <Choice>
        <Name>Version_Choice_UNKNOWN</Name>
        <Description>VersionD_Choice_ALL</Description>
        <Value>Version_UNKNOWN</Value>
        <ExtensionPoint />
    </Choice>
</Choices>
'@

$versionChoiceXml64=@'
<Choices>
    <Choice>
        <Name>Version_Choice_WIN8RTM</Name>
        <Description>VersionD_Choice_ALL</Description>
        <Value>Version_WIN8RTM</Value>
        <ExtensionPoint />
    </Choice>
    <Choice>
        <Name>Version_Choice_WIN7RTM</Name>
        <Description>VersionD_Choice_ALL</Description>
        <Value>Version_WIN7RTM</Value>
        <ExtensionPoint>
            <Default />
        </ExtensionPoint>
    </Choice>
    <Choice>
        <Name>Version_Choice_WINVISTA2</Name>
        <Description>VersionD_Choice_ALL</Description>
        <Value>Version_WINVISTA2</Value>
        <ExtensionPoint />
    </Choice>
    <Choice>
        <Name>Version_Choice_UNKNOWN</Name>
        <Description>VersionD_Choice_ALL</Description>
        <Value>Version_WINVISTA2</Value>
        <ExtensionPoint />
    </Choice>
</Choices>
'@

$displayChoiceXml = @'
<Choices>
    <Choice>
        <Name>DisplayN_Choice_256COLOR</Name>
        <Description>DisplayD_Choice_256COLOR</Description>
        <Value>Display_256COLOR</Value>
        <ExtensionPoint />
    </Choice>
    <Choice>
        <Name>DisplayN_Choice_16BITCOLOR</Name>
        <Description>DisplayD_Choice_16BITCOLOR</Description>
        <Value>Display_16BITCOLOR</Value>
        <ExtensionPoint />
    </Choice>
    <Choice>
        <Name>DisplayN_Choice_640x480</Name>
        <Description>DisplayD_Choice_640x480</Description>
        <Value>Display_640x480</Value>
        <ExtensionPoint />
    </Choice>
    <Choice>
        <Name>DisplayN_Choice_HIGHDPIAWARE</Name>
        <Description>DisplayD_Choice_HIGHDPIAWARE</Description>
        <Value>Display_HIGHDPIAWARE</Value>
        <ExtensionPoint />
    </Choice>
    <Choice>
        <Name>DisplayN_Choice_UNKNOWN</Name>
        <Description>DisplayD_Choice_UNKNOWN</Description>
        <Value>Display_UNKNOWN</Value>
        <ExtensionPoint />
    </Choice>
</Choices>
'@

$typeDefinition = @"

using System;
using System.Collections;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;


public class WerUtil
{
    const int MAX_PATH = 260;
    const int GPLK_USER = 0x00000001;
    const int MAX_LAYER_LENGTH = 256;
    const int WCHAR_SIZE = 2;
    const uint LAYER_APPLIED_FROM_WIZARD = 0x00000010;
    const uint LAYER_APPLIED_FROM_WIZ_CLOUD = 0x00000020;
    const uint SCS_64BIT_BINARY = 6;

    [DllImport("pcwutl.dll", EntryPoint="SendPcwWerReport", CharSet=CharSet.Unicode)]
    [return : MarshalAs(UnmanagedType.Bool)]
    public static extern bool SendPcwWerReport(String ExePath, bool FixesWorked, String ResultFile, String MatchingInfoFile);    

    [DllImport("pcwutl.dll", EntryPoint="GetTempFile", CharSet=CharSet.Unicode)]
    [return : MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetTempFile(String Prefix, StringBuilder ResultFilePath);

    public static String GetTempFilePath()
    {
        StringBuilder resultPath = new StringBuilder(MAX_PATH);
        if (GetTempFile("PCW", resultPath))
        {
            return resultPath.ToString();
        }

        return String.Empty;
    }

    [DllImport("pcwutl.dll", EntryPoint="GetMatchingInfo", CharSet=CharSet.Unicode)]
    [return : MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetMatchingInfo(String ExePath, StringBuilder OutputPath);

    public static String GetMatchingFileInfo(String ExePath)
    {
        StringBuilder resultPath = new StringBuilder(MAX_PATH);
        if (GetMatchingInfo(ExePath, resultPath))
        {
            return resultPath.ToString();
        }

        return String.Empty;
    }

    [DllImport("pcwutl.dll", EntryPoint="LogAeEvent", CharSet=CharSet.Unicode)]
    [return : MarshalAs(UnmanagedType.Bool)]
    public static extern bool LogAeEvent(String ExecutablePath, String CompatibilityLayer, uint ScenarioType, bool FixWorked, String FileIdStr, String ProgramIdStr);

    public static String GetMediaType(String ExePath)
    {
        DriveInfo driveInfo = new DriveInfo(Path.GetPathRoot(ExePath));
        return driveInfo.DriveType.ToString();
    }

    [DllImport("pcwutl.dll", EntryPoint="RetrieveFileAndProgramId", CharSet=CharSet.Unicode)]
    public static extern void RetrieveFileAndProgramId(String ExePath, StringBuilder FileId, StringBuilder ProgramId);

    public static ArrayList MapFilePathToId(String ExePath)
    {
        StringBuilder fileId = new StringBuilder(MAX_PATH);
        StringBuilder programId = new StringBuilder(MAX_PATH);

        RetrieveFileAndProgramId(ExePath, fileId, programId);

        ArrayList idInfo = new ArrayList();
        idInfo.Add(fileId.ToString());
        idInfo.Add(programId.ToString());

        return idInfo;
    }

    [DllImport("apphelp.dll", EntryPoint="SetPermLayerState", CharSet=CharSet.Unicode)]
    [return : MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetPermLayerState(String wszPath, String wszLayer, uint dwFlags, bool bMachine, bool bEnable);

    public static void ApplyCompatMode(String ExePath, ArrayList LayersToApply, ArrayList LayersToRemove, bool CloudLayer)
    {   
        uint Flag;

        if (CloudLayer)
        {
            Flag = LAYER_APPLIED_FROM_WIZ_CLOUD;
        }
        else
        {
            Flag = LAYER_APPLIED_FROM_WIZARD;
        }
        
        foreach (Object layer in LayersToApply)
        {
            if ((String)layer != String.Empty)
            {
                SetPermLayerState(ExePath, (String)layer, Flag, false, true);
            }
        }
        foreach (Object layer in LayersToRemove)
        {
            if ((String)layer != String.Empty)
            {
                SetPermLayerState(ExePath, (String)layer, 0, false, false);
            }
        }
    }

    [DllImport("apphelp.dll", EntryPoint="SdbGetPermLayerKeys", CharSet=CharSet.Unicode)]
    [return : MarshalAs(UnmanagedType.Bool)]
    public static extern bool SdbGetPermLayerKeys(String pwszPath, StringBuilder pwszLayers, out uint pdwBytes, uint dwFlags);

    public static String GetExistingCompatMode(String ExePath)
    {
        StringBuilder existingLayers = new StringBuilder(MAX_LAYER_LENGTH);
        uint existingLayersSize = (uint)existingLayers.Capacity*WCHAR_SIZE;

        if (SdbGetPermLayerKeys(ExePath, existingLayers, out existingLayersSize, GPLK_USER))
        {
            return existingLayers.ToString();
        }

        return String.Empty;
    }

    [DllImport("apphelp.dll", EntryPoint="SdbSetPermLayerKeys", CharSet=CharSet.Unicode)]
    [return : MarshalAs(UnmanagedType.Bool)]
    public static extern bool SdbSetPermLayerKeys(String wszPath, String wszLayers, bool bMachine);

    public static void OverwriteCompatMode(String ExePath, String ModeToApply)
    {
        SdbSetPermLayerKeys(ExePath, ModeToApply, false);
    }

    public static String EscapePath(String Path)
    {
        if (Path == null)
        {
            return null;
        }
        return Path.Replace("$", "`$");
    }   

    [DllImport("kernel32.dll", EntryPoint="GetBinaryTypeW", CharSet=CharSet.Unicode)]
    [return : MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetBinaryType(string lpApplicationName, out uint lpBinaryType);

    public static bool AppIs64Bit(String AppPath)
    {
        uint binaryType;

        if (GetBinaryType(AppPath, out binaryType))
        {
            if (binaryType == SCS_64BIT_BINARY)
            {
                return true;
            }
        }

        return false;
    }

    [DllImport("pcwutl.dll", EntryPoint="GetLayerFromGenome", CharSet=CharSet.Unicode)]
    [return : MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetLayerFromGenome(String ExePath, StringBuilder Layer, out bool IsConfidentVistaPlus);

    public static Array GetGenomeLayer(String ExePath)
    {
        bool IsConfidentVistaPlus = false;
        StringBuilder resultPath = new StringBuilder(MAX_PATH);
        Array retArray = Array.CreateInstance(typeof(Object), 2);
        if (GetLayerFromGenome(ExePath, resultPath, out IsConfidentVistaPlus))
        {
            retArray.SetValue(resultPath.ToString(), 0);            
        }
        else
        {
            retArray.SetValue(String.Empty, 0);
        }
        retArray.SetValue(IsConfidentVistaPlus, 1);

        return retArray;
    }

    [DllImport("pcwutl.dll", EntryPoint = "LogPCWDebugEvent", CharSet = CharSet.Unicode)]
    public static extern void LogPCWDebugEvent(string DebugString, Int64 qwDebugValue);


    public static void LogDebugEvent(String Message, Int64 DebugValue)
    {
        LogPCWDebugEvent(Message, DebugValue);
    }
    
    [DllImport("pcwutl.dll", EntryPoint = "SendSQMForTSRun", CharSet = CharSet.Unicode)]
    public static extern bool SendSQMForTSRun(uint LaunchedFrom, uint COSResponse, string COSLayers, string AppliedLayers, uint UserAction, uint NotUsed, uint ProblemFixed);


    public static void SendSqmForSession(uint LaunchedFrom, uint COSResponse, string COSLayers, string AppliedLayers, uint UserAction, uint NotUsed, uint ProblemFixed)
    {
        SendSQMForTSRun(LaunchedFrom, COSResponse, COSLayers, AppliedLayers, UserAction, NotUsed, ProblemFixed);
        LogPCWDebugEvent(String.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}", LaunchedFrom, COSResponse, COSLayers, AppliedLayers, UserAction, NotUsed, ProblemFixed), 0);
    }
    
    [DllImport("user32.dll", EntryPoint = "MessageBox", CharSet = CharSet.Unicode)]
    public static extern int MessageBox(uint hwnd, string Message, string Caption, uint Buttons);

    public static void ShowMessage(string Message)
    {
        MessageBox(0, Message, "Debug", 0);
    }  
}
"@

function set-selected([System.Collections.Hashtable]$choice, [bool]$select)
{
    if($select -and -not($choice.ContainsKey("ExtensionPoint")))
    {
        $choice.Add("ExtensionPoint", "<Default />")
    }
    elseif(-not($select) -and ($choice["ExtensionPoint"] -ne $null))
    {
        $choice.Remove("ExtensionPoint")
    }
}

#Function to mark a compatibility mode for addition/removal
function SetCompatMode([string]$compatMode = $(throw $CompatibilityStrings.Throw_NO_MODE), [bool]$apply)
{
    if($apply)
    {
        if(-not($layersToApply.Contains($compatMode)) -and -not($originalCompatMode -match $compatMode))
        {
            $layersToApply.Add($compatMode)
        }
        if($layersToRemove.Contains($compatMode))
        {
            $layersToRemove.Remove($compatMode)
        }
    }
    else
    {
        if(-not($layersToRemove.Contains($compatMode)))
        {
            $layersToRemove.Add($compatMode)
        }
        if($layersToApply.Contains($compatMode))
        {
            $layersToApply.Remove($compatMode)
        }
    }
}

#Function to determine the problem(s) the user is having.
function GetProblemSelection()
{
    #Select the xml appropriate for the binary type, and parse it into a list of choices
    #as expected by the diag framework
    $choices = New-Object System.Collections.ArrayList

    if(([System.IO.Path]::GetExtension($targetPath) -eq ".msi"))
	{
        $choiceDoc = [xml] $problemChoiceXmlMsi
    }
    else
    {
        $choiceDoc = [xml] $problemChoiceXml
    }
    
    #The following code segment is used multiple times to generate "dynamic choices"
    #with default selection based on the layers already stored in the registry.
    #Consolidating the code into one function would be preferable, but unfortunately
    #this causes problems with the global persistence of the "choices" variable - I'm
    #not sure why at the moment.

    $choiceDoc.SelectNodes("Choices/Choice") | foreach {
        $choice = @{}
        foreach ($node in $_.ChildNodes)
        {
            if($node.InnerXml -ne [string]::Empty)
            {
                $choice.Add($node.Name, $node.InnerXml)
            }
        }

        #localize the name and description
        $key = $choice["Name"]
        $choice["Name"] = $CompatibilityStrings.$key
        $key = $choice["Description"]
        $choice["Description"] = $CompatibilityStrings.$key

        $choices += $choice
    }

    #Determine if a version layer is set
    if(($AllVersionModes.Split(' ') | where {$originalCompatMode -match $_}) -ne $null)
    {
        $choices | where {$_["Value"] -eq "VersionProblem"} | foreach {
            set-selected $_ $true
        }
    }
    else
    {
        $choices | where {$_["Value"] -eq "VersionProblem"} | foreach {
            set-selected $_ $false
        }
    }

    #Determine if a display layer is set. Vista or better has a different selection
    #than XP or earlier

    if ($isVistaPlus)
    {
        $DisplayModes = $VistaPlusDisplayMode
    }
    else
    {
        $DisplayModes = $AllDisplayModes
    }

    if(($DisplayModes.Split(' ') | where {$originalCompatMode -match $_}) -ne $null)
    {
        $choices | where {$_["Value"] -eq "DisplayProblem"} | foreach {
            set-selected $_ $true
        }
    }
    else
    {
        $choices | where {$_["Value"] -eq "DisplayProblem"} | foreach {
            set-selected $_ $false
        }
    }

    #Determine if the runasadmin layer is set
    if($originalCompatMode -match $RunAsAdminCompatMode)
    {
        $choices | where {$_["Value"] -eq "RunAsAdminProblem"} | foreach {
            set-selected $_ $true
        }
    }
    else
    {
        $choices | where {$_["Value"] -eq "RunAsAdminProblem"} | foreach {
            set-selected $_ $false
        }
    }

    $problemChoices = Get-DiagInput -id IT_ProblemDisplay -choice $choices

    $mask = 0
    foreach($selection in $problemChoices)
    {
        if($selection -eq "VersionProblem")
        {
            $mask = $mask -bor $VersionProblem
        }
        
        if($selection -eq "DisplayProblem")
        {
            $mask = $mask -bor $DisplayProblem
        }
        
        if($selection -eq "RunAsAdminProblem")
        {
            $mask = $mask -bor $RunAsAdminProblem
        }
        
        if($selection -eq "UnknownProblem" -and ($problemChoices.Length -eq 1))
        {
            $mask = $mask -bor $VersionProblem

            if(-not($appIs64Bit)) 
            {
                $mask = $mask -bor $DisplayProblem
            }

            $mask = $mask -bor $RunAsAdminProblem
        }
    }

    Set-Variable -name problemMask -value $mask -scope global
}

#Function to determine the user's version choice
#Unlike other problem categories, this is a single selection
function GetVersionLayer([bool]$showInteraction)
{
    $choices = New-Object System.Collections.ArrayList
    $choiceDoc = $null

    if($appIs64Bit)
    {
        $choiceDoc = [xml] $versionChoiceXml64
    }
    else
    {
        $choiceDoc = [xml] $versionChoiceXml
    }

    $choiceDoc.SelectNodes("Choices/Choice") | foreach { 
        $choice = @{} 
        foreach ($node in $_.ChildNodes)
        { 
            if($node.InnerXml -ne [string]::Empty) 
            { 
                $choice.Add($node.Name, $node.InnerXml) 
            } 
        }

        #localize the name and description
        $key = $choice["Name"]
        $choice["Name"] = $CompatibilityStrings.$key
        $key = $choice["Description"]
        $choice["Description"] = $CompatibilityStrings.$key

        $choices += $choice 
    }

    $choices | where {$originalCompatMode -match $CompatibilityModes[$_["Value"]]} | foreach {
        set-selected $_ $true
    }
    $choices | where {-not($originalCompatMode -match $CompatibilityModes[$_["Value"]])} | foreach {
        set-selected $_ $false
    }

    $versionChoice = $null

    if($showInteraction)
    {
        $versionChoice = Get-DiagInput -id IT_WindowsVersions -choice $choices
        if(($versionChoice -ne [String]::Empty) -and ($versionChoice -ne $null))
        {
            Set-Variable -name solutionSelected -value $true -scope global
            SetCompatMode $CompatibilityModes[$versionChoice] $true
        }
    }

    #Make sure unselected choices are not set
    foreach($choice in $choices)
    {
        if($versionChoice -eq $null -or -not($CompatibilityModes[$choice["Value"]] -eq $CompatibilityModes[$versionChoice]) -and ($originalCompatMode -match $CompatibilityModes[$choice["Value"]]))
        {
            SetCompatMode $CompatibilityModes[$choice["Value"]] $false
        }    
    }
}

#Function to determine the user's display choices
#We allow the user to select multiple symptoms
function GetDisplayLayers([bool]$showInteraction)
{
    #pull xml nodes from string declared above, parse it into array of dictionary objects
    $choices = New-Object System.Collections.ArrayList
    $choiceDoc = [xml] $displayChoiceXml

    $choiceDoc.SelectNodes("Choices/Choice") | foreach {
        $choice = @{} 
        foreach ($node in $_.ChildNodes)
        { 
            if($node.InnerXml -ne [string]::Empty) 
            { 
                $choice.Add($node.Name, $node.InnerXml) 
            } 
        }

        #localize the name and description for the item
        $key = $choice["Name"]
        $choice["Name"] = $CompatibilityStrings.$key
        $key = $choice["Description"]
        $choice["Description"] = $CompatibilityStrings.$key

        $choices += $choice 
    }

    $choices | where {$CompatibilityModes.ContainsKey($_["Value"]) -and ($originalCompatMode -match $CompatibilityModes[$_["Value"]])} | foreach {
        set-selected $_ $true
    }
    $choices | where {-not($CompatibilityModes.ContainsKey($_["Value"]) -and ($originalCompatMode -match $CompatibilityModes[$_["Value"]]))} | foreach {
        set-selected $_ $false
    }

    $displayChoices = New-Object System.Collections.ArrayList
    $selectionFound = $false

    if($showInteraction)
    {
        $displayChoices = Get-DiagInput -id IT_DisplayProblems -choice $choices

        foreach($selection in $displayChoices)
        {
            $selectionFound = $true

            if(($selection -ne "Display_UNKNOWN") -and ($selection -ne [String]::Empty))
            {
                Set-Variable -name solutionSelected -value $true -scope global
                SetCompatMode $CompatibilityModes[$selection] $true
            }
        }
    }

    #Make sure unselected choices are not set
    foreach($choice in $choices)
    {
        $choiceIsSelected = $false
        foreach($selectedChoice in $displayChoices)
        {
            if($selectedChoice -eq $choice["Value"])
            {
                $choiceIsSelected = $true
                break
            }
        }

        if(-not($choiceIsSelected) -and ($originalCompatMode -match $CompatibilityModes[$choice["Value"]]))
        {
            SetCompatMode $CompatibilityModes[$choice["Value"]] $false
        }
    }
}

#Function to set the text for the summary page.
function GetSummary()
{

    set-variable compatModeParam $CompatibilityStrings.Version_Choice_DEFAULT -scope global
    set-variable displayModeParam $CompatibilityStrings.Display_Choice_DEFAULT -scope global
    set-variable accessModeParam $CompatibilityStrings.Access_Choice_DEFAULT -scope global

    foreach($layer in $layersToApply)
    {
        
        if($allVersionModes -match $layer)
        {
             set-variable compatModeParam $CompatibilityModeStrings[$layer] -scope global
        }

        if($allDisplayModes -match $layer)
        {
            $displayMode = Get-Variable -name displayModeParam -valueOnly -scope global
            if($displayMode -eq $CompatibilityStrings.Display_Choice_DEFAULT)
            {
                $displayMode = [String]::Empty
            }
            if($displayMode -ne [String]::Empty)
            {
                $displayMode += $displaySpacer
            }
            $displayMode += $CompatibilityModeStrings[$layer]

            Set-Variable -name displayModeParam -value $displayMode -scope global
        }
        
        if($RunAsAdminCompatMode -eq $layer)
        {
            Set-Variable -name accessModeParam -value $CompatibilityStrings.Access_Choice_ADMIN -scope global
        }

        if($MsiAutoCompatMode -eq $layer)
        {
             set-variable compatModeParam $CompatibilityModeStrings[$layer] -scope global
        }
    }

    $originalCompatMode.Split($delimiters.ToCharArray()) | foreach {
        if(-not($_ -eq [String]::Empty) -and -not($layersToRemove.Contains($_)))
        {
            if($AllVersionModes.Contains($_))
            {
                set-variable compatModeParam $CompatibilityModeStrings[$_] -scope global
            }

            if($AllDisplayModes.Contains($_))
            {
                $displayMode = Get-Variable -name displayModeParam -valueOnly -scope global
                if($displayMode -eq $CompatibilityStrings.Display_Choice_DEFAULT)
                {
                    $displayMode = [String]::Empty
                }
                if (-not($displayMode.Contains($CompatibilityModeStrings[$_])))
                {
                    if($displayMode -ne [String]::Empty)
                    {
                        $displayMode += $displaySpacer
                    }
                    $displayMode += $CompatibilityModeStrings[$_]
                }
                Set-Variable -name displayModeParam -value $displayMode -scope global
            }
        
            if($RunAsAdminCompatMode -eq $_)
            {
               Set-Variable -name accessModeParam -value $CompatibilityStrings.Access_Choice_ADMIN -scope global
            }

            if($MsiAutoCompatMode -eq $_)
            {
                set-variable compatModeParam $CompatibilityModeStrings[$_] -scope global
            }
        }
    }
}

#alters the contents of $Env:RecommendedLayer based on if the exe is an executable or if it is a Vista+ era app
function FilterRecommendedLayers([bool]$isExecutable, [bool]$isVistaPlus) {

    [string]$tempLayers = ""
    [string]$PreVistaDisplayMode = "256COLOR 16BITCOLOR 640X480"

    $Env:RecommendedLayer.Split(' ') | foreach {
    
        $Layer = $_

        #Executables don't get the MSIAUTO layer
        if ($isExecutable -and ($Layer -eq "MSIAUTO")) {
            $Layer = ""
        }

        #Vista+ era apps don't get some display layers
        if ($isVistaPlus -and ($PreVistaDisplayMode.Contains($Layer))) {
            $Layer = ""
        }

        $tempLayers = "{0} {1}" -f $Layer, $tempLayers
    }

    $tempLayers = $tempLayers.Trim(' ')
    
    if ($tempLayers -eq "") {
        $Env:RecommendedLayer = "NONE"
    } else {
        $Env:RecommendedLayer = $tempLayers
    }
}

# This block of code sets up the manual troubleshooting portion.
#

$werUtilType = Add-Type -TypeDefinition $typeDefinition -PassThru -IgnoreWarnings

$targetPath = $werUtilType::EscapePath($targetPath)

if($targetPath -eq $null)
{
    throw $CompatibilityStrings.Throw_INVALID_PATH
}

# Initialize

set-variable verifyResponse "Verify_TRYAGAIN" -scope global
set-variable solutionSelected $false -scope global
set-variable appIs64Bit $false -scope global
set-variable tsChoice "ts_MANUAL" -scope global
set-variable RecLayerIsCloudLayer $false -scope global
set-variable isVistaPlus $false -scope global

$autoFix = $true
$isExecutable = ([System.IO.Path]::GetExtension($targetPath) -eq ".exe")
$layersToApply = New-Object System.Collections.ArrayList
$layersToRemove = New-Object System.Collections.ArrayList

if($werUtilType::AppIs64Bit($targetPath))
{
    set-variable appIs64Bit $true -scope global
}

# This section will show the screen that asks if you want to do the recommended thing or the manual troubleshooting

if($isExecutable)
{
    $tsChoice = Get-DiagInput -id IT_AutoTroubleshoot
}

# if they chose the recommended layer or the app is an MSI, we will pick a recommended layer for them.
# recommended layers come from (in order of precedence)
# 1. COS recommendation 
# 2. Genome recommendation
# 3. hardcoded value (WIN8RTM)
#

#Query app genome for genome layer and if the app is known Vista+
$genomeArray = $werUtilType::GetGenomeLayer($targetPath)
Set-Variable -name isVistaPlus -value $genomeArray[1] -scope global

FilterRecommendedLayers $isExecutable $isVistaPlus

if ($Env:RecommendedLayer -eq "NONE")
{
    $Env:RecommendedLayer = $genomeArray[0]
    if($Env:RecommendedLayer -eq [String]::Empty)
    {
        $Env:RecommendedLayer = "WIN8RTM"
    }
}
else
{
    Set-Variable -name RecLayerIsCloudLayer $true -scope global
}

# If they chose the recommended layer, we do some processing to remove applied stuff
if(($tsChoice -eq "ts_AUTO") -or -not($isExecutable))
{
    #For each layer in the recommended layer string (usually one, can be more)
    #we will add it to the layersToApply. We will remove all the other layers
    #from the app.

    if($isExecutable)
    {
        $Env:RecommendedLayer.Split(' ') | foreach {
            if ($AllTSLayers.Contains($_) -and (-not($layersToApply.Contains($_))))
            {
                $layersToApply.Add($_)
            }
        }
    }
    else
    {
        #MSIs also get the RunAsAdmin layer
        $layersToApply.Add($RunAsAdminCompatMode)
        $layersToApply.Add("MSIAUTO")
    }

    $AllTSLayers.Split(' ') | foreach {
        if (-not($layersToApply.Contains($_))) 
        {
            $layersToRemove.Add($_)
        }
    }

    Set-Variable -name solutionSelected -value $true -scope global
}
else
{
    $autoFix = $false
}

try
{
    #Loop until either the problem is solved or the user gives up
    do
    {

        #Get the original layers string
        $originalCompatMode = $werUtilType::GetExistingCompatMode($targetPath)
    
        if(-not($autoFix))
        {
            #Reset variables
            set-variable solutionSelected $false -scope global
            set-variable problemMask 0 -scope global
            $layersToApply.Clear()
            $layersToRemove.Clear()
        
            #Ask the user to identify their symptoms if this is an exe
            if($isExecutable)
            {
                GetProblemSelection
            }
            #MSIs get the RunAsAdmin layer automatically applied and automatically
            #prompt the user to select a version layer
            else
            {
                $mask = $VersionProblem
                $mask = $mask -bor $RunAsAdminProblem
                set-Variable problemMask $mask -scope global
                set-Variable solutionSelected $true -scope global
            }
    
            [int]$mask = get-Variable -name problemMask -valueOnly -scope global
        
            if($mask -band $VersionProblem)
            {
                GetVersionLayer($true)    
            }
            else
            {
                GetVersionLayer($false)
            }

            if($mask -band $DisplayProblem)
            {

                #If this is a vista era or later application, the only display problem
                #we attempt to fix will be the high DPI aware issue
                if ($isVistaPlus)
                {

                    #No interaction for this; just need to set the layer
                    if(-not($originalCompatMode -match $VistaPlusDisplayMode))
                    {
                        SetCompatMode $VistaPlusDisplayMode $true                                
                    }
                    set-Variable solutionSelected $true -scope global 

                }
                else
                {
                    GetDisplayLayers($true)
                }
            }
            else
            {
                GetDisplayLayers($false)
            }            
                                    
            if($mask -band $RunAsAdminProblem)
            {
                #No interaction for this; just need to set the layer
                if(-not($originalCompatMode -match $RunAsAdminCompatMode))
                {
                    SetCompatMode $RunAsAdminCompatMode $true
                }
                set-Variable solutionSelected $true -scope global
            }
            else
            {
                if($originalCompatMode -match $RunAsAdminCompatMode)
                {
                    SetCompatMode $RunAsAdminCompatMode $false
                }
            }
        }
        
        if($solutionSelected)
        {
            $quotedPath = "`""+$targetPath+"`""

            #Get the command line of the scheduled task we'll run
            $schedTaskCmd = "{0}\rundll32.exe {0}\pcwutl.dll,LaunchApplication `"{1}`"" -f [System.Environment]::SystemDirectory,$quotedPath
        
            #Make the registry entry
            $werUtilType::ApplyCompatMode($targetPath, $layersToApply, $layersToRemove, $RecLayerIsCloudLayer)
        
            GetSummary
            $param1 = Get-Variable -name compatModeParam -valueOnly -scope global
            $param2 = Get-Variable -name displayModeParam -valueOnly -scope global
            $param3 = Get-Variable -name accessModeParam -valueOnly -scope global
        
            if($isExecutable)
            { 
                $getDiagCmd = "Get-DiagInput -id IT_Summary -parameter @{ `"ExePath`"=`"$schedTaskCmd`";"
        
                $title = $CompatibilityStrings.Text_AppName_Title
                $getDiagCmd+="`"AppName`"=`"$title $appName`";"

                if ($param1 -notlike "*None*")
                {
                   $param1 += "`n";
                   $title = $CompatibilityStrings.Text_Version_Title
                   $getDiagCmd+="`"CompatMode`"=`"$title $param1`";"
                }
                else
                {
                   $getDiagCmd+="`"CompatMode`"=`"`";"
                }
                if ($param2 -notlike "*Normal*")
                {
                   $param2 += "`n";
                   $title = $CompatibilityStrings.Text_Display_Title
                   $getDiagCmd+="`"DisplayMode`"=`"$title  $param2`";"
                   $displayWarning = $CompatibilityStrings.Text_Display_Warning
                   $getDiagCmd+="`"DisplayWarning`"=`"`n$displayWarning `n`";"
                }
                else
                {
                   $getDiagCmd+="`"DisplayMode`"=`"`";"
                   $getDiagCmd+="`"DisplayWarning`"=`"`";"
                }
                if ($param3 -notlike "*Normal*")
                {
                   $title = $CompatibilityStrings.Text_Access_Title
                   $getDiagCmd+="`"AccessMode`"=`"$title  $param3`";"
                }
                else
                {
                   $getDiagCmd+="`"AccessMode`"=`"`";"
                }
        
                $getDiagCmd+="}"
            
                $runResult = Invoke-Expression $getDiagCmd              
    
                while($runResult -ne "Run")
                {
                    $getDiagCmd = $getDiagCmd.Replace("IT_Summary ", "IT_Summary_Error ")
                    $runResult = Invoke-Expression $getDiagCmd
                }            
            }
            else
            {
                $getDiagCmd = "Get-DiagInput -id IT_SummaryMSI -parameter @{ `"ExePath`"=`"$schedTaskCmd`"; `"CompatMode`"=`"$param1`" }"
                $runResult = Invoke-Expression $getDiagCmd
        
                while($runResult -ne "Run")
                {
                    $getDiagCmd = $getDiagCmd.Replace("IT_SummaryMSI ", "IT_SummaryMSI_Error ")
                    $runResult = Invoke-Expression $getDiagCmd
                }
            }
         
            #Determine if we need to re-run the wizard
            if($isExecutable)
            {
                $verifyResponse = Get-DiagInput -id IT_VerifySolution
            }
            else
            {
                $verifyResponse = Get-DiagInput -id IT_MsiVerifySolution
            }
        
            Set-Variable -name launchError -value $false -scope global
        
            #Remove applied settings if the user decided the problem wasn't fixed.
            if($verifyResponse -ne "Verify_YES")
            {
                $werUtilType::OverwriteCompatMode($targetPath, $originalCompatMode)
            }
        }
        else
        {
            $verifyResponse = Get-DiagInput -id IT_NoSolution
        }

        if ($verifyResponse -eq "Verify_TRYAGAIN") {
            $autoFix = $false
            $RecLayerIsCloudLayer = $false
        }
    }
    while($verifyResponse -eq "Verify_TRYAGAIN")

    #Send SQM about the TS session

    #Applied Layers in a space separated string
    $Env:SQMAppliedLayers = ""
   
    $layersToApply | foreach {
    
        $Env:SQMAppliedLayers = "{0} {1}" -f $_, $Env:SQMAppliedLayers
        $Env:SQMAppliedLayers = $Env:SQMAppliedLayers.TrimEnd(' ')
    }

    #Need to preserve this for SQM
    #Recommended layer (2) or did they manually troubleshoot (3) ?
    if ($autoFix) {
        $Env:SQMUserAction = 2
    } else {
        $Env:SQMUserAction = 3
    }

    #Was the problem solved?
    if ($verifyResponse -eq "Verify_YES")
    {
        #If the user said this fixed their problem, we note it so we can show 'fixed' at the end of the process (in the vf_ script)   
        $Env:AppFixed = $true
        $Env:SQMProblemSolved = 1
    } else {
        $Env:SQMProblemSolved = 0
    }

    $werUtilType::SendSqmForSession($Env:SQMLaunchMethod, $Env:SQMCOSResponse, $Env:SQMCOSLayers, $Env:SQMAppliedLayers, $Env:SQMUserAction, 0, $Env:SQMProblemSolved)  

    #Clear the settings and exit without sending a report.
    if($verifyResponse -eq "Verify_UNDO")
    {
        $layersToApply.Clear()
        $layersToRemove.Clear()
    
        $AllVersionModes.Split(' ') | foreach {
        $layersToRemove.Add($_)
        }   
    
        $AllDisplayModes.Split(' ') | foreach {
        $layersToRemove.Add($_)
        }
        
        $layersToRemove.Add($RunAsAdminCompatMode)
    
        $werUtilType::ApplyCompatMode($targetPath, $layersToApply, $layersToRemove, $false)
        exit
    }

    Write-DiagProgress -activity $CompatibilityStrings.Text_Activity_SAVING -status $CompatibilityStrings.Text_Status_GENERATING
    
    #Get final state of compat mode
    $finalCompatMode = $werUtilType::GetExistingCompatMode($targetPath)

    #Get the applied compatibility modes as user-friendly strings
    $modesApplied = New-Object System.Collections.ArrayList
    foreach($layerApplied in $layersToApply) {
        if ($RunAsAdminCompatMode -eq $layerApplied)
        {
            $modesApplied.Add($CompatibilityStrings.Access_Choice_ADMIN)
        }
        else
        {
            $modesApplied.Add($CompatibilityModeStrings[$layerApplied])
        }
    } 
    
    #Push the compat mode settings into the diag report
    $fixesWorked = $false
    $fixVerified = $CompatibilityStrings.Text_Report_SolutionNo
    if ($verifyResponse -eq "Verify_YES")
    {
        $fixesWorked = $true
        $fixVerified = $CompatibilityStrings.Text_Report_SolutionYes
    }

    $modesApplied.Add($fixVerified)  #Add fix verification to modes applied because Update-DiagReport only seems to accept one update

    $modesApplied | convertto-xml | Update-DiagReport -id compatMode -name $CompatibilityStrings.Text_Report_CompatName -description $CompatibilityStrings.Text_Report_CompatDesc -verbosity Informational
    
    #Generate XML containing wizard results
    $resultFile = $werUtilType::GetTempFilePath()
    
    $mediaType = $werUtilType::GetMediaType($targetPath)
    
    $compatResult = $resultFailure
    if ($verifyResponse -eq "Verify_YES")
    {
        $compatResult = $resultSuccess
    }
    
    $wizardResultXml = @"
<?xml version="1.0" encoding="UTF-16" ?> 
<CompatWizardResults ApplicationName="$targetPath"
                     ApplicationPath="$targetPath"
                     MediaType="$mediaType" 
                     CompatibilityResult="$compatResult"> 
</CompatWizardResults>
"@
    
    $resultXml = [xml] $wizardResultXml
    
    #Create a new LAYER element for each layer applied
    foreach($layerApplied in $layersToApply)
    {
        $element = $resultXml.CreateElement("LAYER")
        $attribute = $resultXml.CreateAttribute("NAME")
        $attribute.set_Value($layerApplied)
        $element.SetAttributeNode($attribute)
        $resultXml.CompatWizardResults.AppendChild($element)
    }
    
    $resultXml.save($resultFile)
    
    #Get the matching file information
    $matchingInfoFile = $werUtilType::GetMatchingFileInfo($targetPath)
    
    #Send WER report
    $result = $werUtilType::SendPcwWerReport($targetPath, $fixesWorked, $resultFile, $matchingInfoFile)
    
    #Delete the temporary files
    Remove-Item -path $resultFile -erroraction silentlycontinue
    Remove-Item -path $matchingInfoFile -erroraction silentlycontinue
    
    #Write event to event log
    $fileIdInfo = $werUtilType::MapFilePathToId($targetPath)
    $result = $werUtilType::LogAeEvent($targetPath, $finalCompatMode, 203, $fixesWorked, $fileIdInfo[0], $fileIdInfo[1])    
}

# Handle cancel event
finally
{
    if(($originalCompatMode -ne $null) -and ($verifyResponse -eq "Verify_TRYAGAIN"))
    {
        $werUtilType::OverwriteCompatMode($targetPath, $originalCompatMode)
    }
}
