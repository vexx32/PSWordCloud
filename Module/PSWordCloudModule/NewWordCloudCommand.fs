namespace PSWordCloud
open System
open System.Collections.Generic
open System.Linq
open System.Management.Automation
open System.Numerics
open System.Text.RegularExpressions
open System.Threading.Tasks
open PSWordCloud.Extensions
open PSWordCloud.Utils
open SkiaSharp


[<Cmdlet(VerbsCommon.New, "WordCloud"); Alias("nwc","wcloud","newcloud")>]
type NewWordCloudCommand() =
    inherit PSCmdlet()

    let mutable _resolvedPath = String.Empty

    [<Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "ColorBackground")>]
    [<Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "ColorBackground-Mono")>]
    [<Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "FileBackground")>]
    [<Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "FileBackground-Mono")>]
    [<Alias("InputString", "Text", "String", "Words", "Document", "Page")>]
    [<AllowEmptyString()>]
    member val public InputObject : PSObject = null
        with get, set

    [<Parameter(Mandatory = true, Position = 0, ParameterSetName = "ColorBackground")>]
    [<Parameter(Mandatory = true, Position = 0, ParameterSetName = "ColorBackground-Mono")>]
    [<Parameter(Mandatory = true, Position = 0, ParameterSetName = "FileBackground")>]
    [<Parameter(Mandatory = true, Position = 0, ParameterSetName = "FileBackground-Mono")>]
    [<Alias("OutFile", "ExportPath", "ImagePath")>]
    member public self.Path
        with get() = _resolvedPath
        and set(v) = _resolvedPath <- self.SessionState.Path.GetUnresolvedProviderPathFromPSPath(v)

    [<Parameter(ParameterSetName = "ColorBackground")>]
    [<Parameter(ParameterSetName = "ColorBackground-Mono")>]
    //[<ArgumentCompleter(typeof<ImageSizeCompleter>)>]
    //[<TransformToSKSizeI()>]
    member val public ImageSize = SKSizeI(3840, 2160)
        with get, set


