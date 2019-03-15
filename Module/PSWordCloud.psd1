#
# Module manifest for module 'PSWordCloud'
#
# Generated by: Joel Sallow
#
# Generated on: 7/8/2018
#

@{

    # Script module or binary module file associated with this manifest.
    RootModule           = 'PSWordCloud.psm1'

    # Version number of this module.
    ModuleVersion        = '2.0.0'

    # Supported PSEditions
    CompatiblePSEditions = @('Core')

    # ID used to uniquely identify this module
    GUID                 = 'c63b9cfe-cca8-40a6-9002-f19555c036b7'

    # Author of this module
    Author               = 'Joel Sallow'

    # Company or vendor of this module
    CompanyName          = 'None'

    # Copyright statement for this module
    Copyright            = '(c) 2019 Joel Sallow (/u/ta11ow, @vexx32). All rights reserved.'

    # Description of the functionality provided by this module
    Description          = 'Turn your scripts and documents into pretty and practical word clouds!'

    # Minimum version of the Windows PowerShell engine required by this module
    PowerShellVersion    = '5.1'

    # Name of the Windows PowerShell host required by this module
    # PowerShellHostName = ''

    # Minimum version of the Windows PowerShell host required by this module
    # PowerShellHostVersion = ''

    # Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
    # DotNetFrameworkVersion = ''

    # Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
    # CLRVersion = ''

    # Processor architecture (None, X86, Amd64) required by this module
    # ProcessorArchitecture = @()

    # Modules that must be imported into the global environment prior to importing this module
    # RequiredModules       = @()

    # Assemblies that must be loaded prior to importing this module
    # RequiredAssemblies = @()

    # Script files (.ps1) that are run in the caller's environment prior to importing this module.
    # ScriptsToProcess      = @()

    # Type files (.ps1xml) to be loaded when importing this module
    # TypesToProcess = @()

    # Format files (.ps1xml) to be loaded when importing this module
    # FormatsToProcess = @()

    # Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
    # NestedModules = @()

    # Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
    FunctionsToExport    = @()

    # Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
    CmdletsToExport      = @('New-WordCloud')

    # Variables to export from this module
    VariablesToExport    = '*'

    # Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
    AliasesToExport      = @('wordcloud', 'wcloud', "nwc")

    # DSC resources to export from this module
    # DscResourcesToExport = @()

    # List of all modules packaged with this module
    # ModuleList = @()

    # List of all files packaged with this module
    FileList             = @(
        'PSWordCloud.psd1'
        'PSWordCloud.psm1'
        'PSWordCloudCmdlet.dll'
        'linux-x64/libSkiaSharp.so'
        'linux-x64/SkiaSharp.dll'
        'osx/libSkiaSharp.dylib'
        'osx/SkiaSharp.dll'
        'win-x64/libSkiaSharp.dll'
        'win-x64/SkiaSharp.dll'
        'win-x86/libSkiaSharp.dll'
        'win-x86/SkiaSharp.dll'
    )

    # Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
    PrivateData          = @{

        PSData = @{

            # Tags applied to this module. These help with module discovery in online galleries.
            Tags       = @('Graphic', 'Art', 'wordcloud', 'generator', 'image')

            # A URL to the license for this module.
            LicenseUri = 'https://github.com/vexx32/PSWordCloud/blob/master/LICENSE'

            # A URL to the main website for this project.
            ProjectUri = 'https://github.com/vexx32/PSWordCloud'

            # A URL to an icon representing this module.
            # IconUri = ''

            # ReleaseNotes of this module
            # ReleaseNotes = ''

        } # End of PSData hashtable

    } # End of PrivateData hashtable

    # HelpInfo URI of this module
    # HelpInfoURI = ''

    # Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
    # DefaultCommandPrefix = ''
}

