namespace PSWordCloud

open System
open System.Collections
open System.Collections.Generic
open System.Collections.ObjectModel
open System.Linq
open System.Management.Automation
open System.Reflection
open System.Runtime.CompilerServices
open SkiaSharp
open System.Management.Automation

module Operators =
    let inline (!>) (x:^a) : ^b = ((^a or ^b) : (static member op_Implicit : ^a -> ^b) x)

type internal WordOrientation =
    | Horizontal = 0
    | Vertical = 1
    | FlippedVertical = 2


module Utils =
    let FocusWordScale = 1.3f
    let BleedAreaScale = 1.15f
    let MinSaturation = 5.0f
    let MinBrightness = 25.0f
    let MaxPercentWidth = 0.75f
    let PaddingBaseScale = 0.05f
    let StrokeBaseScale = 0.02f

    let FontManager = SKFontManager.Default

    let FontList =
        FontManager.FontFamilies.OrderBy((fun x -> x), StringComparer.OrdinalIgnoreCase)

    let ColorLibrary =
        typeof<SKColor>.GetFields (BindingFlags.Static ||| BindingFlags.Public)
        |> Seq.map (fun field -> (field.Name, field.GetValue(null) :?> SKColor))
        |> Map.ofSeq

    let ColorNames =
        ColorLibrary
        |> Seq.map (fun item -> item.Key)

    let StandardColors =
        ColorLibrary
        |> Seq.map (fun item -> item.Value)

    let NamedColor color = ColorLibrary.[color]

    let StandardImageSizes =
        [
            ("480x800", ("Mobile Screen Size (small)", SKSizeI(480, 800)))
            ("640x1146", ("Mobile Screen Size (medium)", SKSizeI(640, 1146)))
            ("720p", ("Standard HD 1280x720", SKSizeI(1280, 720)))
            ("1080p", ("Full HD 1920x1080", SKSizeI(1920, 1080)))
            ("4K", ("Ultra HD 3840x2160", SKSizeI(3840, 2160)))
            ("A4", ("816x1056", SKSizeI(816, 1056) ))
            ("Poster11x17", ("1056x1632", SKSizeI(1056, 1632)))
            ("Poster18x24", ("1728x2304", SKSizeI(1728, 2304)))
            ("Poster24x36", ("2304x3456", SKSizeI(2304, 3456)))
        ] |> Map.ofSeq

module Extensions =
    type SKPoint with
        member self.Multiply factor = SKPoint(self.X * factor, self.Y * factor)

    type Single with
        member self.ToRadians() = self * (single Math.PI) / 180.0f

    type Random with
        member self.Shuffle<'T> (array : 'T[]) =
            let mutable n : int = array.Length
            while (n > 1) do
                let k = self.Next(n)
                n <- n - 1
                let temp = array.[n]
                array.[n] <- array.[k]
                array.[k] <- temp

    type SKRect with
        member self.FallsOutside (region : SKRegion) =
            let bounds = SKRect(single region.Bounds.Left, single region.Bounds.Top, single region.Bounds.Right, single region.Bounds.Bottom)
            self.Top < bounds.Top
            || self.Bottom > bounds.Bottom
            || self.Left < bounds.Left
            || self.Right > bounds.Right

    type SKPaint with
        member self.NextWord color wordSize strokeWidth =
            self.TextSize <- wordSize
            self.IsStroke <- false
            self.Style <- SKPaintStyle.StrokeAndFill
            self.StrokeWidth <- wordSize * strokeWidth * Utils.StrokeBaseScale
            self.IsVerticalText <- false
            self.Color = color

        member self.DefaultWord = self.NextWord SKColors.Black

    type SKColor with
        member self.SortValue adjustmentValue =
            let (hue, saturation, brightness) = self.ToHsv()
            let vibranceAdjustment = brightness * (adjustmentValue * 0.5f) / (1.0f - saturation)
            brightness + vibranceAdjustment

    type SKRegion with
        member self.Op(path : SKPath, operation : SKRegionOperation) =
            use pathRegion = new SKRegion()
            pathRegion.SetPath(path) |> ignore
            self.Op(pathRegion, operation)

        member self.Intersects(rect : SKRect) =
            match self.Bounds.IsEmpty with
            | true ->
                use region = new SKRegion()
                region.SetRect(SKRectI.Round(rect)) |> ignore
                self.Intersects(region)
            | x -> x

        member self.Intersects(path : SKPath) =
            match self.Bounds.IsEmpty with
            | true ->
                use region = new SKRegion()
                region.SetPath(path) |> ignore
                self.Intersects(region)
            | x -> x

    type IEnumerable<'T> with
        member self.GetValue key =
            match self with
            | :? PSMemberInfoCollection<PSPropertyInfo> as prop -> prop.[key].Value
            | :? IDictionary as dict -> dict.[key]
            | _ ->
                let memberType = typeof<PSMemberInfoCollection<PSPropertyInfo>>
                let dictType = typeof<IDictionary<_, _>>
                raise (ArgumentException(String.Format("GetValue method only accepts {0} or {1}", memberType, dictType)))