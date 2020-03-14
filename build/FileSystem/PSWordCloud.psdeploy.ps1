Deploy Module {
    By PSGalleryModule {
        FromSource "$env:BUILD_ARTIFACTSTAGINGDIRECTORY/PSWordCloud"
        To FileSystem
        WithOptions @{
            ApiKey = 'FileSystem'
        }
    }
}
