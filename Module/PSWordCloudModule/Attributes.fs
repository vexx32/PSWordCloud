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

type ImageSizeCompleter() =
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

    override self.Transform(intrinsics, inputData) =
        let sideLength = 0
        match inputData with
        | :? SKSize as sz -> sz :> _
        | :? SKSizeI as szI -> szI :> _
        | :? string as s
            ->
            if StandardImageSizes.ContainsKey(s) then
                StandardImageSizes.[s] :> _
            else
                let sizePattern = @"^(?<Width>[\d\.,]+)x(?<Height>[\d\.,]+)(px)?$"
                let numberPattern = @"^(?<SideLength>[\d\.,]+)(px)?$"

                let matchSize = Regex.Match(s, sizePattern)
                if matchSize.Success then
                    SKSizeI(int matchSize.Groups.["Width"], int matchSize.Groups.["Height"])
                else
                    let matchNumber = Regex.Match(s, numberPattern)
                    if matchNumber.Success then
                        let x = int matchNumber.Groups.["SideLength"]
                        SKSizeI(x, x)

        | :? array as arr ->

        | _ ->

(*public override object Transform(EngineIntrinsics engineIntrinsics, object inputData)
        {
            int sideLength = 0;
            switch (inputData)
            {
                case SKSize sk:
                    return sk;
                case short sh:
                    sideLength = sh;
                    break;
                case ushort us:
                    sideLength = us;
                    break;
                case int i:
                    sideLength = i;
                    break;
                case uint u:
                    if (u <= int.MaxValue)
                    {
                        sideLength = (int)u;
                    }

                    break;
                case long l:
                    if (l <= int.MaxValue)
                    {
                        sideLength = (int)l;
                    }

                    break;
                case ulong ul:
                    if (ul <= int.MaxValue)
                    {
                        sideLength = (int)ul;
                    }
                    break;
                case decimal d:
                    if (d <= int.MaxValue)
                    {
                        sideLength = (int)Math.Round(d);
                    }
                    break;
                case float f:
                    if (f <= int.MaxValue)
                    {
                        sideLength = (int)Math.Round(f);
                    }
                    break;
                case double d:
                    if (d <= int.MaxValue)
                    {
                        sideLength = (int)Math.Round(d);
                    }
                    break;
                case string s:
                    if (WCUtils.StandardImageSizes.ContainsKey(s))
                    {
                        return WCUtils.StandardImageSizes[s].Size;
                    }
                    else
                    {
                        var matchWH = Regex.Match(s, @"^(?<Width>[\d\.,]+)x(?<Height>[\d\.,]+)(px)?$");
                        if (matchWH.Success)
                        {
                            try
                            {
                                var width = int.Parse(matchWH.Groups["Width"].Value);
                                var height = int.Parse(matchWH.Groups["Height"].Value);

                                return new SKSizeI(width, height);
                            }
                            catch (Exception e)
                            {
                                throw new ArgumentTransformationMetadataException(
                                    "Could not parse input string as a float value", e);
                            }
                        }

                        var matchSide = Regex.Match(s, @"^(?<SideLength>[\d\.,]+)(px)?$");
                        if (matchSide.Success)
                        {
                            sideLength = int.Parse(matchSide.Groups["SideLength"].Value);
                        }
                    }

                    break;
                case object o:
                    IEnumerable properties = null;
                    if (o is Hashtable ht)
                    {
                        properties = ht;
                    }
                    else
                    {
                        properties = PSObject.AsPSObject(o).Properties;
                    }

                    if (properties.GetValue("Width") != null && properties.GetValue("Height") != null)
                    {
                        // If these conversions fail, the exception will cause the transform to fail.
                        var width = LanguagePrimitives.ConvertTo<int>(properties.GetValue("Width"));
                        var height = LanguagePrimitives.ConvertTo<int>(properties.GetValue("Height"));

                        return new SKSizeI(width, height);
                    }

                    break;
            }

            if (sideLength > 0)
            {
                return new SKSizeI(sideLength, sideLength);
            }

            throw new ArgumentTransformationMetadataException();
        }*)