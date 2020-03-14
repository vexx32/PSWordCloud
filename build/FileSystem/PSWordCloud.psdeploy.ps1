Deploy Module {
    By PSGalleryModule {
        FromSource "$env:BUILD_ARTIFACT_STAGING_DIRECTORY/PSWordCloud"
        To FileSystem
        WithOptions @{
            ApiKey = 'FileSystem'
        }
    }
}
