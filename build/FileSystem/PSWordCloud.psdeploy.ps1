Deploy Module {
    By PSGalleryModule {
        FromSource "$PSScriptRoot/../PSWordCloud"
        To FileSystem
        WithOptions @{
            ApiKey = 'FileSystem'
        }
    }
}
