using System;
using System.Collections;
using System.Collections.Generic;
using System.Management.Automation;
using System.Management.Automation.Language;

namespace PSWordCloud.Completers
{
    public class ImageSizeCompleter : IArgumentCompleter
    {
        public IEnumerable<CompletionResult> CompleteArgument(
            string commandName,
            string parameterName,
            string wordToComplete,
            CommandAst commandAst,
            IDictionary fakeboundParameters)
        {
            foreach (var result in Utils.StandardImageSizes)
            {
                if (string.IsNullOrEmpty(wordToComplete) ||
                    result.Key.StartsWith(wordToComplete, StringComparison.OrdinalIgnoreCase))
                {
                    yield return new CompletionResult(
                        result.Key,
                        result.Key,
                        CompletionResultType.ParameterValue,
                        result.Value.Tooltip);
                }
            }
        }
    }
}
