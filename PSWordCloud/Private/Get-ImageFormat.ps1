function Get-ImageFormat {
    [CmdletBinding()]
    [OutputType([Imaging.ImageFormat])]
    param(
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [ValidateSet("Bmp", "Emf", "Exif", "Gif", "Jpeg", "Png", "Tiff", "Wmf")]
        [string]
        $Format
    )
    process {
        switch ($Format) {
            "Bmp"  { [Imaging.ImageFormat]::Bmp  }
            "Emf"  { [Imaging.ImageFormat]::Emf  }
            "Exif" { [Imaging.ImageFormat]::Exif }
            "Gif"  { [Imaging.ImageFormat]::Gif  }
            "Jpeg" { [Imaging.ImageFormat]::Jpeg }
            "Png"  { [Imaging.ImageFormat]::Png  }
            "Tiff" { [Imaging.ImageFormat]::Tiff }
            "Wmf"  { [Imaging.ImageFormat]::Wmf  }
        }
    }
}