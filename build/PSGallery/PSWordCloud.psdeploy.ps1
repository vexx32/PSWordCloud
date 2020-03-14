Deploy Module {
    By PSGalleryModule {
        FromSource "$PSScriptRoot/../PSWordCloud"
        To PSGallery
        WithOptions @{
            ApiKey = $env:NugetApiKey
        }
    }
}
