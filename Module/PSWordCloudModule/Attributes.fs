namespace PSWordCloud

open System
open System.Collections
open System.Collections.Generic
open System.Collections.ObjectModel
open System.Linq
open System.Management.Automation
open System.Management.Automation.Language
open System.Text.RegularExpressions
open PSWordCloud.Utils
open SkiaSharp
open System.Management.Automation

type SKSizeICompleter() =
    interface IArgumentCompleter with
        member self.CompleteArgument(commandName, parameterName, wordToComplete, commandAst, fakeBoundParameters) =
            seq {
                for item in StandardImageSizes do
                    if (String.IsNullOrEmpty(wordToComplete)
                        || item.Key.StartsWith(wordToComplete, StringComparison.OrdinalIgnoreCase)) then
                        let (tooltip, _) = item.Value
                        yield CompletionResult(item.Key, item.Key, CompletionResultType.ParameterValue, tooltip)
            }

type TransformToSKSizeIAttribute() =
    inherit ArgumentTransformationAttribute()

    override self.Transform(intrinsics, input) =
        let sideLength = 0
        match input with
        | :? SKSize as sz -> sz :> obj
        | :? SKSizeI as szI -> szI :> obj
        | :? string as s ->
            if StandardImageSizes.ContainsKey(s) then
                StandardImageSizes.[s] :> obj
            else
                let sizePattern = @"^(?<Width>[\d\.,]+)x(?<Height>[\d\.,]+)(px)?$"
                let numberPattern = @"^(?<SideLength>[\d\.,]+)(px)?$"
                let matchSize = Regex.Match(s, sizePattern)
                if matchSize.Success then
                    SKSizeI(int matchSize.Groups.["Width"].Value, int matchSize.Groups.["Height"].Value) :> obj
                else
                    let matchNumber = Regex.Match(s, numberPattern)
                    if matchNumber.Success then
                        let x = int matchNumber.Groups.["SideLength"].Value
                        SKSizeI(x, x) :> obj
                    else
                        raise (ArgumentTransformationMetadataException())
        | x ->
            let properties =
                match x with
                | :? IDictionary as x -> x :> IEnumerable
                | x -> PSObject.AsPSObject(x).Properties :> IEnumerable

            SKSizeI(ValueFrom properties "Width" |> To<int>, ValueFrom properties "Width" |> To<int>) :> obj

type TypefaceCompleter() =
    interface IArgumentCompleter with
        member self.CompleteArgument(commandName, parameterName, wordToComplete, commandAst, fakeBoundParameters) =
            seq {
                let target = wordToComplete.TrimStart('"').TrimEnd('"')
                for item in FontList do
                    if String.IsNullOrEmpty(wordToComplete)
                        || item.StartsWith(target, StringComparison.OrdinalIgnoreCase)
                    then
                        if item.Contains ' '
                            || item.Contains '#'
                            || wordToComplete.StartsWith '"'
                        then
                            let result = String.Format("\"{0}\"", item)
                            yield CompletionResult(result, item, CompletionResultType.ParameterValue, item)
                        else
                            yield CompletionResult(item, item, CompletionResultType.ParameterValue, item)
            }

type TransformToSKTypefaceAttribute() =
    inherit ArgumentTransformationAttribute()

    override self.Transform(intrinsics, input) =
        match input with
        | :? SKTypeface as x -> x :> obj
        | :? string as s -> FontManager.MatchFamily(s, SKFontStyle.Normal) :> obj
        | x ->
            let properties =
                match x with
                | :? IDictionary as d -> d :> IEnumerable
                | p -> PSObject.AsPSObject(p).Properties :> IEnumerable

            let style =
                let weight = ValueFrom properties "Weight" |> As<SKFontStyleWeight>
                let width = ValueFrom properties "Width" |> As<SKFontStyleWidth>
                let slant = ValueFrom properties "Slant" |> As<SKFontStyleSlant>

                match (weight, width, slant) with
                | (Some x, Some y, Some z) -> new SKFontStyle(x, y, z)
                | (None, Some y, Some z) -> new SKFontStyle(SKFontStyleWeight.Normal, y, z)
                | (Some x, None, Some z) -> new SKFontStyle(x, SKFontStyleWidth.Normal, z)
                | (Some x, Some y, None) -> new SKFontStyle(x, y, SKFontStyleSlant.Upright)
                | (Some x, None, None) -> new SKFontStyle(x, SKFontStyleWidth.Normal, SKFontStyleSlant.Upright)
                | (None, Some y, None) -> new SKFontStyle(SKFontStyleWeight.Normal, y, SKFontStyleSlant.Upright)
                | (None, None, Some z) -> new SKFontStyle(SKFontStyleWeight.Normal, SKFontStyleWidth.Normal, z)
                | (None, None, None) -> SKFontStyle.Normal

            let familyName = ValueFrom properties "FamilyName" |> To<string>
            FontManager.MatchFamily(familyName, style) :> obj

type SKColorCompleter() =
    interface IArgumentCompleter with
        member self.CompleteArgument(commandName, parameterName, wordToComplete, commandAst, fakeBoundParameters) =
            seq {
                for color in ColorNames do
                    if String.IsNullOrEmpty(wordToComplete)
                        || color.StartsWith(wordToComplete, StringComparison.OrdinalIgnoreCase)
                    then
                        let colorValue = ColorLibrary.[color]
                        let tooltip = String.Format("{0} (R: {1}, G: {2}, B: {3}, A: {4})", color, colorValue.Red, colorValue.Green, colorValue.Blue, colorValue.Alpha)
                        yield CompletionResult(color, color, CompletionResultType.ParameterValue, tooltip)
            }

type TransformToSKColorAttribute() =
    inherit ArgumentTransformationAttribute()

    let matchColor name =
        seq {
            match name with
            | knownColor when ColorNames.Contains(knownColor) -> yield ColorLibrary.[knownColor]
            | clear when String.Equals(clear, "transparent", StringComparison.OrdinalIgnoreCase) -> yield SKColor.Empty
            | patternString when WildcardPattern.ContainsWildcardCharacters(patternString) ->
                let pattern = WildcardPattern(patternString, WildcardOptions.IgnoreCase)
                for color in ColorNames do
                    if pattern.IsMatch(color) then yield ColorLibrary.[color]
            | x ->
                let (success, color) = SKColor.TryParse(x)
                if success then yield color
        }

    let transformObject (input : obj) =
        let values =
            match input with
            | :? Array as arr -> arr
            | x -> Seq.toArray [ x ] :> Array
        seq {
            for item in values do
                match item with
                | :? string as s -> for color in matchColor s do yield color
                | x ->
                    let properties =
                        match x with
                        | :? IDictionary as d -> d :> IEnumerable
                        | p -> PSObject.AsPSObject(p).Properties :> IEnumerable
                    let red = ValueFrom properties "Red" |> As<byte>
                    let green = ValueFrom properties "Green" |> As<byte>
                    let blue = ValueFrom properties "Blue" |> As<byte>
                    let alpha = ValueFrom properties "Alpha" |> As<byte>

                    yield
                        match (red, green, blue, alpha) with
                        | (None, None, None, None) -> SKColors.Transparent
                        | _ -> SKColor(red |? 0uy, green |? 0uy, blue |? 0uy, alpha|? 255uy)
        }

    override self.Transform(intrinsics, data) =
        let results =
            match data with
            | :? SKColor as c -> seq { yield c }
            | :? string as s -> matchColor s
            | x -> transformObject x
            |> Seq.toList

        match results with
        | unit when unit.Length = 1 -> unit.[0] :> obj
        | x -> x :> obj
