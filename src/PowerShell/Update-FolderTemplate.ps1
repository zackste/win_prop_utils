<#
.SYNOPSIS
    Updates the folder template settings for a specified folder type in Windows Explorer.

.DESCRIPTION
    This script allows you to update the folder template settings for various folder types such as Generic, Documents,
    Pictures, Music, and Videos. You can specify prioritized columns, sort columns, optionally backup existing
    settings, and optionally reset all template folders.

.LINK
    https://codeplexarchive.org/project/prop
        The Property System Command Line Interface (prop.exe) is a command line tool for managing the Windows Property
        System. It can be used to identify properties names to use as columns with this script as well as retrieve
        detailed information about each property.
#>
[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'UpdateTemplate')]
param(
    # The known folder type to update. Default is 'Music'.
    # Valid options are: Generic, Documents, Pictures, Music, Videos.
    [ValidateSet('Generic', 'Documents', 'Pictures', 'Music', 'Videos')]
    [Parameter(Mandatory = $true, ParameterSetName = 'UpdateTemplate')]
    [Parameter(Mandatory = $true, ParameterSetName = 'ShowDefaultColumns')]
    [string]$Template,

    # A hashtable containing prioritized columns for the folder template. Values in entry '0' will be prioritized
    # displayed in order by default. Values in entry '1' will be shown in the list of default optional columns that
    # can be added to the view. For example:
    # @{
    #     '0' = @('System.ItemNameDisplay', 'System.DateModified', 'System.ItemTypeText')
    #     '1' = @('System.Size', 'System.Media.Duration', 'System.Audio.ChannelCount')
    # }
    # If not specified, the script will use the default columns for the specified template type.
    [Parameter(Mandatory = $false, ParameterSetName = 'UpdateTemplate')]
    [hashtable]$PrioritizedColumns = $null,

    # An optional array of columns to sort by. If specified, these will be set as the default sort order for the folder
    # template. Default is $null, meaning no change to the existing sort order.
    [Parameter(Mandatory = $false, ParameterSetName = 'UpdateTemplate')]
    [string[]]$SortColumns = $null,

    # If specified, the existing column settings will be backed up before applying the new settings. This should only
    # be run on the first run of the script to avoid overwriting existing backups.
    [Parameter(Mandatory = $false, ParameterSetName = 'UpdateTemplate')]
    [switch]$BackupTemplate,

    # If specified, all customized column settings for all folders assigned the specified template type will be reset
    # to the default settings. This will remove any customizations made to each folder after the template was applied.
    [Parameter(Mandatory = $false, ParameterSetName = 'UpdateTemplate')]
    [switch]$ResetAllTemplateFolders,

    # If specified, the script will show the default columns for the specified template type.
    [Parameter(ParameterSetName = 'ShowDefaultColumns', Mandatory = $true)]
    [switch]$ShowDefaultColumns,

    [Parameter(ParameterSetName = 'ListAvailableColumns', Mandatory = $true)]
    [switch]$ListAvailableColumns
)

$defaultColumns = @{
    'Music' = @{
        '0' = @(
            'System.ItemNameDisplay'
            'System.DateModified'
            'System.ItemTypeText'
            'System.Size'
            'System.Media.Duration'
            'System.Audio.ChannelCount'
            'System.Audio.SampleSize'
            'System.Audio.SampleRate'
            'System.Music.TrackNumber'
        )
        '1' = @(
            'System.Title'
            'System.Music.Artist'
            'System.Music.AlbumTitle'
            'System.DateCreated'
            'System.Music.AlbumArtist'
            'System.Music.Genre'
            'System.DRM.IsProtected'
            'System.Rating'
            'System.Media.Year'
        )
    }
}

if ($ShowDefaultColumns) {
    # Show default columns for the specified template
    if ($defaultColumns.ContainsKey($Template)) {
        $columns = $defaultColumns[$Template]
        Write-Host "Default ordered columns for template '$Template':"
        $columns['0'] | ForEach-Object { [PSCustomObject]@{CanonicalName = $_; Order = 'Primary' } }
        $columns['1'] | ForEach-Object { [PSCustomObject]@{CanonicalName = $_; Order = 'Optional' } }
    }
    else {
        Write-Error "No default columns found for template '$Template'."
    }
    return
}

if (-not $PrioritizedColumns -and -not $ListAvailableColumns) {
    # Use default columns if none specified
    $PrioritizedColumns = $defaultColumns[$Template]
    if ($PrioritizedColumns) {
        Write-Verbose "Using default columns for template '$Template':"
        Write-Verbose "Primary Columns: $($PrioritizedColumns | ConvertTo-Json -Depth 3 | Out-String)"
    }
    else {
        Write-Error "No default columns found for template '$Template'. Please specify the PrioritizedColumns parameter."
        return
    }
}

$PROPDESC_ENUMFILTER = @{
    "PDEF_ALL"             = 0
    "PDEF_SYSTEM"          = 1
    "PDEF_NONSYSTEM"       = 2
    "PDEF_VIEWABLE"        = 3
    "PDEF_QUERYABLE"       = 4
    "PDEF_INFULLTEXTQUERY" = 5
    "PDEF_COLUMN"          = 6
}

$csharpNativeHooks = @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public class UpdateFolderTemplateNativeHelper
{
    [Guid("6f79d558-3e96-4549-a1d1-7d75d2288814")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    public partial interface IPropertyDescription
    {
        IntPtr GetPropertyKey();
        [return: MarshalAs(UnmanagedType.LPWStr)] string GetCanonicalName();
        ushort GetPropertyType();
        [return: MarshalAs(UnmanagedType.LPWStr)] string GetDisplayName();
        [return: MarshalAs(UnmanagedType.LPWStr)] string GetEditInvitation();
        IntPtr GetTypeFlags(IntPtr mask);
        IntPtr GetViewFlags();
        uint GetDefaultColumnWidth();
        IntPtr GetDisplayType();
        uint GetColumnState();
        IntPtr GetGroupingRange();
        IntPtr GetRelativeDescriptionType();
        [PreserveSig] int GetRelativeDescription(IntPtr propvar1, IntPtr propvar2, [MarshalAs(UnmanagedType.LPWStr)] out string ppszDesc1, [MarshalAs(UnmanagedType.LPWStr)] out string ppszDesc2);
        IntPtr GetSortDescription();
        [return: MarshalAs(UnmanagedType.LPWStr)] string GetSortDescriptionLabel([MarshalAs(UnmanagedType.Bool)] bool fDescending);
        IntPtr GetAggregationType();
        [PreserveSig] int GetConditionType(out int pcontype, out IntPtr popDefault);
        [PreserveSig] int GetEnumTypeList(Guid riid, out IntPtr ppv);
        [PreserveSig] int CoerceToCanonicalValue(ref IntPtr ppropvar);
        [return: MarshalAs(UnmanagedType.LPWStr)] string FormatForDisplay(IntPtr propvar, int pdfFlags);
        [PreserveSig] int IsValueCanonical(IntPtr propvar);
    }

    [Guid("1f9fc1d0-c39b-4b26-817f-011967d3440e")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    public partial interface IPropertyDescriptionList
    {
        uint GetCount();
        IPropertyDescription GetAt(uint iElem, ref Guid riid);
    }

    [DllImport("Propsys.dll", CallingConvention = CallingConvention.StdCall)]
    private static extern int PSEnumeratePropertyDescriptions(
        int filterOn,
        in Guid riid,
        out IPropertyDescriptionList ppv);

    // https://learn.microsoft.com/en-us/windows/win32/seccrypto/common-hresult-values
    private const int E_FAIL = unchecked((int)0x80004005);

    public static string GetDisplayName(IPropertyDescription propertyDescription)
    {
        try
        {
            return propertyDescription.GetDisplayName();
        }
        catch (COMException come) when (come.ErrorCode == E_FAIL)
        {
            return null;
        }
    }

    public static IEnumerable<string[]> GetAvailableColumns(int filterOn)
    {
        Guid guid = typeof(IPropertyDescriptionList).GUID;
        int result = PSEnumeratePropertyDescriptions(filterOn, guid, out IPropertyDescriptionList list);
        Marshal.ThrowExceptionForHR(result);
        var size = (int)list.GetCount();
        for (int i = 0; i < size; i++)
        {
            guid = typeof(IPropertyDescription).GUID;
            var item = list.GetAt((uint)i, ref guid);
            yield return new string[] { item.GetCanonicalName(), GetDisplayName(item)};
        }
    }
}
"@

# List available properties filtered to ones that can be used as columns.
if ($ListAvailableColumns) {
    try {
        Add-Type -TypeDefinition $csharpNativeHooks -Language CSharp
    }
    catch {
        Write-Error "Failed to load native hooks for property enumeration: $_"
        # return 1
    }
    try {
        $availableColumns = [UpdateFolderTemplateNativeHelper]::GetAvailableColumns($PROPDESC_ENUMFILTER.PDEF_COLUMN)
        Write-Host "Available columns for folder templates:"
        $availableColumns | ForEach-Object {[PSCustomObject]@{CanonicalName = $_[0]; DisplayName = $_[1]}}
        # foreach ($column in $availableColumns) {
        #     Write-Output $column
        # }
    }
    catch {
        Write-Error "Failed to enumerate available columns: $_"
        return 1
    }
    return
}

$folderTypeIdMap = @{
    'Generic'   = '5c4f28b5-f869-4e84-8e60-f11db97c5cc7'
    'Documents' = '7d49d726-3c21-4f05-99aa-fdc2c9474656'
    'Pictures'  = 'b3690e58-e961-423b-b687-386ebfd83239'
    'Music'     = '94d6ddcc-4a68-4175-a374-bd584a510b78'
    'Videos'    = '5fa96407-7e77-483c-ac93-691d05850de8'
}

$backupValueNames = @{
    'ColumnList' = 'ColumnList.bak'
    'SortByList' = 'SortByList.bak'
}

$resolvedTemplate = $folderTypeIdMap[$Template]
if (-not $resolvedTemplate) {
    Write-Error "Invalid template specified. Valid options are: Generic, Documents, Pictures, Music, Videos."
    return
}

$resolvedKey = Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes\{$resolvedTemplate}\TopViews\{00000000-0000-0000-0000-000000000000}"
if (-not $resolvedKey) {
    Write-Error "Folder template for $Template not found in registry."
    return
}

$columnListEntries = @() + ($PrioritizedColumns['0'] | ForEach-Object { "0$_" }) + ($PrioritizedColumns['1'] | ForEach-Object { "1$_" })
$resolvedColumnList = "prop:$($columnListEntries -join ';')"
if ($PSCmdlet.ShouldProcess("Update folder template for $Template with columns: $resolvedColumnList", "Set column list and sort order")) {
    Write-Verbose "Updating folder template for $Template with columns: $resolvedColumnList"
}

if ($resolvedKey.Property -contains 'ColumnList' -and ($BackupTemplate -or $null -eq $resolvedKey.GetValue($backupValueNames['ColumnList'], $null))) {
    # Backup first
    $existingColumnList = $resolvedKey | Get-ItemProperty -Name ColumnList -ErrorAction SilentlyContinue | ForEach-Object ColumnList
    $resolvedKey | Set-ItemProperty -Name ($backupValueNames['ColumnList']) -Value $existingColumnList -Force
}

$resolvedKey | Set-ItemProperty -Name ColumnList -Value $resolvedColumnList

if ($SortColumns -and $SortColumns.Count -gt 0) {
    # if ($BackupTemplate -and $resolvedKey.Property -contains 'SortByList') {
    if ($resolvedKey.Property -contains 'SortByList' -and ($BackupTemplate -or $null -eq $resolvedKey.GetValue($backupValueNames['SortByList'], $null))) {
        # Backup first
        $existingSortByList = $resolvedKey | Get-ItemProperty -Name SortByList -ErrorAction SilentlyContinue | ForEach-Object SortByList
        $resolvedKey | Set-ItemProperty -Name ($backupValueNames['SortByList']) -Value $existingSortByList. -Force
    }

    $sortByList = "prop:$($SortColumns -join ';')"
    if ($PSCmdlet.ShouldProcess("Update folder template for $Template with sort columns: $sortByList", "Set sort order")) {
        Write-Verbose "Updating folder template for $Template with sort columns: $sortByList"
    }
    $resolvedKey | Set-ItemProperty -Name SortByList -Value $sortByList
}

if ($ResetAllTemplateFolders) {
    # Reset all existing column settings for the specified template
    Get-ChildItem "HKCU:\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags\*\Shell\{$resolvedTemplate}" | Remove-ItemProperty -Name ColInfo, Sort -ErrorAction SilentlyContinue
}

Write-Host "Windows Explorer may need to be restarted or a Windows logout and login may be needed for changes to take effect."