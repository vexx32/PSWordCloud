using System.Collections;
using System.Management.Automation;
using SkiaSharp;

namespace PSWordCloud.Attributes
{
    public class TransformToSKTypefaceAttribute : ArgumentTransformationAttribute
    {
        public override object Transform(EngineIntrinsics engineIntrinsics, object inputData)
        {
            return inputData switch
            {
                SKTypeface t => t,
                string s => Utils.FontManager.MatchFamily(s, SKFontStyle.Normal),
                _ => CreateTypefaceFromObject(inputData),
            };
        }

        private static SKTypeface CreateTypefaceFromObject(object input)
        {
            IEnumerable properties = input is Hashtable ht
                ? ht
                : PSObject.AsPSObject(input).Properties;

            SKFontStyle style;
            if (properties.GetValue("FontWeight") is not null
                || properties.GetValue("FontSlant") is not null
                || properties.GetValue("FontWidth") is not null)
            {
                object? weightValue = properties.GetValue("FontWeight");
                SKFontStyleWeight weight = weightValue is null
                    ? SKFontStyleWeight.Normal
                    : weightValue.ConvertTo<SKFontStyleWeight>();

                object? slantValue = properties.GetValue("FontSlant");
                SKFontStyleSlant slant = slantValue is null
                    ? SKFontStyleSlant.Upright
                    : slantValue.ConvertTo<SKFontStyleSlant>();

                object? widthValue = properties.GetValue("FontWidth");
                SKFontStyleWidth width = widthValue is null
                    ? SKFontStyleWidth.Normal
                    : widthValue.ConvertTo<SKFontStyleWidth>();

                style = new SKFontStyle(weight, width, slant);
            }
            else
            {
                style = properties.GetValue("FontStyle") is SKFontStyle customStyle
                    ? customStyle
                    : SKFontStyle.Normal;
            }

            string familyName = properties.GetValue("FamilyName")?.ConvertTo<string>() ?? string.Empty;
            return Utils.FontManager.MatchFamily(familyName, style);
        }
    }
}
