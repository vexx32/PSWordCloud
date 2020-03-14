Deploy Module {
    By PSGalleryModule {
        FromSource "$env:BUILD_ARTIFACT_STAGING_DIRECTORY/PSWordCloud"
        To PSGallery
        WithOptions @{
            ApiKey = $env:NugetApiKey
        }
    }
}
