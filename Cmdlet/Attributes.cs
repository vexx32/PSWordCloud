using System;
using System.Collections.Generic;
using System.Management.Automation;
using SkiaSharp;

namespace PSWordCloud
{
    /*
            {
                param(Command, Parameter, WordToComplete, CommandAst, FakeBoundParams)
                if (!WordToComplete) {
                    return [ColorTransformAttribute]::ColorNames
                }
                else {
                    return [ColorTransformAttribute]::ColorNames.Where{ _ -match "^WordToComplete" }
                }
            }
    */
    public class ColorCompleter : IArgumentCompleter
    {

    }

    class SizeTransformAttribute : ArgumentTransformationAttribute {
        static Dictionary<string, Size> StandardSizes = @{
            '720p'  = [Size]::new(1280, 720)
            '1080p' = [Size]::new(1920, 1080)
            '4K'    = [Size]::new(4096, 2160)
        }

        [object] Transform([EngineIntrinsics]engineIntrinsics, [object] inputData) {
            Size = switch (inputData) {
                { _ -is [Size] } {
                    _
                    break
                }
                { _ -is [SizeF] } {
                    _.ToSize()
                    break
                }
                { _ -is [int] -or _ -is [double] } {
                    [Size]::new(_, _)
                    break
                }
                { _ -in [SizeTransformAttribute]::StandardSizes.Keys } {
                    [SizeTransformAttribute]::StandardSizes[_]
                    break
                }
                { _ -is [string] } {
                    if (_ -match '^(?<Width>[\d\.,]+)x(?<Height>[\d\.,]+)(px)?') {
                        [Size]::new(Matches['Width'], Matches['Height'])
                        break
                    }

                    if (_ -match '^(?<Size>[\d\.,]+)(px)?') {
                        [Size]::new(Matches['Size'], Matches['Size'])
                        break
                    }
                }
                default {
                    throw [ArgumentTransformationMetadataException]::new("Unable to convert entered value inputData to a valid [System.Drawing.Size].")
                }
            }

            Area = Size.Height * Size.Width
            if (Area -ge 100 * 100 -and Area -le 20000 * 20000) {
                return Size
            }
            else {
                throw [ArgumentTransformationMetadataException]::new(
                    "Specified size inputData is either too small to use for an image size, or would exceed GDI+ limitations."
                )
            }
        }
    }

    class ColorTransformAttribute : ArgumentTransformationAttribute {
        static [string[]] ColorNames = @(
            [KnownColor].GetEnumNames()
            "Transparent"
        )

        [object] Transform([EngineIntrinsics]engineIntrinsics, [object] inputData) {
            Items = switch (inputData) {
                { _ -eq null -or _ -eq 'Transparent' } {
                    [Color]::Transparent
                    continue
                }
                { _ -as [KnownColor] } {
                    [Color]::FromKnownColor(_ -as [KnownColor])
                    continue
                }
                { _ -is [Color] } {
                    _
                    continue
                }
                { _ -is [string] } {
                    if (_ -match 'R(?<Red>[0-9]{1,3})G(?<Green>[0-9]{1,3})B(?<Blue>[0-9]{1,3})') {
                        [Color]::FromArgb(Matches['Red'], Matches['Green'], Matches['Blue'])
                        continue
                    }

                    if (_ -match 'R(?<Red>[0-9]{1,3})G(?<Green>[0-9]{1,3})B(?<Blue>[0-9]{1,3})A(?<Alpha>[0-9]{1,3})') {
                        [Color]::FromArgb(Matches['Alpha'], Matches['Red'], Matches['Green'], Matches['Blue'])
                        continue
                    }

                    if (MatchingValues = [KnownColor].GetEnumNames() -like _) {
                        (MatchingValues -as [KnownColor[]]).ForEach{ [Color]::FromKnownColor(_) }
                    }
                }
                { _ -is [int] } {
                    [Color]::FromArgb(_)
                    continue
                }
                default {
                    throw [ArgumentTransformationMetadataException]::new("Could not convert value '_' to a valid [System.Drawing.Color] or [System.Drawing.KnownColor].")
                }
            }

            return Items
        }
    }

    class FileTransformAttribute : ArgumentTransformationAttribute {
        [object] Transform([EngineIntrinsics]engineIntrinsics, [object] inputData) {
            Items = switch (inputData) {
                { _ -as [FileInfo] } {
                    _
                    break
                }
                { _ -is [string] } {
                    Path = Resolve-Path -Path _
                    if (@(Path).Count -gt 1) {
                        throw [ArgumentTransformationMetadataException]::new("Multiple files found, please enter only one: (Path -join ', ')")
                    }

                    if (Test-Path -Path Path -PathType Leaf) {
                        [FileInfo]::new(Path)
                    }

                    break
                }
                default {
                    throw [ArgumentTransformationMetadataException]::new("Could not convert value '_' to a valid [System.IO.FileInfo] object.")
                }
            }

            return Items
        }
    }
}