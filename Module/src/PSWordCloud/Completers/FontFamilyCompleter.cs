using System;
using System.Collections;
using System.Collections.Generic;
using System.Management.Automation;
using System.Management.Automation.Language;

namespace PSWordCloud.Completers
{
    public class FontFamilyCompleter : IArgumentCompleter
    {
        public IEnumerable<CompletionResult> CompleteArgument(
            string commandName,
            string parameterName,
            string wordToComplete,
            CommandAst commandAst,
            IDictionary fakeBoundParameters)
        {
            string matchString = wordToComplete.TrimStart('"').TrimEnd('"');
            foreach (string font in Utils.FontList)
            {
                if (string.IsNullOrEmpty(wordToComplete)
                    || font.StartsWith(matchString, StringComparison.OrdinalIgnoreCase))
                {
                    if (font.Contains(' ') || font.Contains('#') || wordToComplete.StartsWith("\""))
                    {
                        var result = string.Format("\"{0}\"", font);
                        yield return new CompletionResult(result, font, CompletionResultType.ParameterName, font);
                    }
                    else
                    {
                        yield return new CompletionResult(font, font, CompletionResultType.ParameterName, font);
                    }
                }
            }
        }
    }
}
