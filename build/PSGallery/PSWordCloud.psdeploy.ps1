Deploy Module {
    By PSGalleryModule {
        FromSource "$env:BUILD_ARTIFACTSTAGINGDIRECTORY/PSWordCloud"
        To PSGallery
        WithOptions @{
            ApiKey = $env:NugetApiKey
        }
    }
}
