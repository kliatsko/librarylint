@{
    # Module manifest for LibraryLint
    # Generated for PowerShell Gallery compatibility

    # Script module or binary module file associated with this manifest
    RootModule = 'LibraryLint.ps1'

    # Version number of this module (updated with each release)
    ModuleVersion = '5.2.3'

    # ID used to uniquely identify this module
    GUID = 'a1b2c3d4-e5f6-7890-abcd-ef1234567890'

    # Author of this module
    Author = 'Nick Kliatsko'

    # Company or vendor of this module
    CompanyName = 'Personal'

    # Copyright statement for this module
    Copyright = '(c) 2026 Nick Kliatsko. All rights reserved.'

    # Description of the functionality provided by this module
    Description = 'A modular toolkit for media library organization, cleanup, and management. Designed for managing movie and TV show collections with support for Kodi/Plex/Jellyfin/Emby-compatible naming and metadata.'

    # Minimum version of the PowerShell engine required by this module
    PowerShellVersion = '5.1'

    # Processor architecture (None, X86, Amd64) required by this module
    # ProcessorArchitecture = ''

    # Modules that must be imported into the global environment prior to importing this module
    # RequiredModules = @()

    # Assemblies that must be loaded prior to importing this module
    # RequiredAssemblies = @()

    # Script files (.ps1) that are run in the caller's environment prior to importing this module
    # ScriptsToProcess = @()

    # Type files (.ps1xml) to be loaded when importing this module
    # TypesToProcess = @()

    # Format files (.ps1xml) to be loaded when importing this module
    # FormatsToProcess = @()

    # Modules to import as nested modules of the module specified in RootModule
    NestedModules = @(
        'modules\Mirror.psm1',
        'modules\Sync.psm1'
    )

    # Functions to export from this module
    FunctionsToExport = @()

    # Cmdlets to export from this module
    CmdletsToExport = @()

    # Variables to export from this module
    VariablesToExport = @()

    # Aliases to export from this module
    AliasesToExport = @()

    # DSC resources to export from this module
    # DscResourcesToExport = @()

    # List of all modules packaged with this module
    # ModuleList = @()

    # List of all files packaged with this module
    FileList = @(
        'LibraryLint.ps1',
        'Run-LibraryLint.bat',
        'README.md',
        'CHANGELOG.md',
        'GETTING_STARTED.md',
        'LICENSE',
        'modules\Mirror.psm1',
        'modules\Sync.psm1',
        'config\config.example.json'
    )

    # Private data to pass to the module specified in RootModule
    PrivateData = @{
        PSData = @{
            # Tags applied to this module for discoverability
            Tags = @(
                'media',
                'library',
                'organization',
                'movies',
                'tv-shows',
                'kodi',
                'plex',
                'jellyfin',
                'emby',
                'metadata',
                'nfo',
                'tmdb',
                'tvdb',
                'subtitles',
                'trailers',
                'cleanup',
                'rename',
                'organize'
            )

            # A URL to the license for this module
            LicenseUri = 'https://github.com/kliatsko/librarylint/blob/main/LICENSE'

            # A URL to the main website for this project
            ProjectUri = 'https://github.com/kliatsko/librarylint'

            # A URL to an icon representing this module
            # IconUri = ''

            # ReleaseNotes of this module
            ReleaseNotes = 'See CHANGELOG.md for detailed release notes.'

            # Prerelease tag (for pre-release versions)
            # Prerelease = ''

            # Flag to indicate whether the module requires explicit user acceptance for install/update/save
            # RequireLicenseAcceptance = $false

            # External dependencies not managed by PowerShell Gallery
            ExternalModuleDependencies = @()
        }
    }

    # HelpInfo URI of this module
    HelpInfoURI = 'https://github.com/kliatsko/librarylint#readme'
}
