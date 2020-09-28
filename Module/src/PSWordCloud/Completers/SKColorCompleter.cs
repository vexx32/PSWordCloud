using System;
using System.Collections;
using System.Collections.Generic;
using System.Management.Automation;
using System.Management.Automation.Language;
using SkiaSharp;

namespace PSWordCloud.Completers
{
    public class SKColorCompleter : IArgumentCompleter
    {
        public IEnumerable<CompletionResult> CompleteArgument(
            string commandName,
            string parameterName,
            string wordToComplete,
            CommandAst commandAst,
            IDictionary fakeBoundParameters)
        {
            foreach (string color in Utils.ColorNames)
            {
                if (string.IsNullOrEmpty(wordToComplete)
                    || color.StartsWith(wordToComplete, StringComparison.OrdinalIgnoreCase))
                {
                    SKColor colorValue = Utils.ColorLibrary[color];
                    yield return new CompletionResult(
                        completionText: color,
                        listItemText: color,
                        CompletionResultType.ParameterValue,
                        toolTip: string.Format(
                            "{0} (R: {1}, G: {2}, B: {3}, A: {4})",
                            color, colorValue.Red, colorValue.Green, colorValue.Blue, colorValue.Alpha));
                }
            }
        }
    }
}
