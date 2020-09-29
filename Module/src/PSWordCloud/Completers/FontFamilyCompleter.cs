using System;
using System.Collections;
using System.Collections.Generic;
using System.Management.Automation;
using System.Management.Automation.Language;
using System.Text.RegularExpressions;

namespace PSWordCloud.Completers
{
    public class FontFamilyCompleter : IArgumentCompleter
    {
        private const string requireQuotesPattern = "[ #|-]";
        private const string requireEscapingPattern = "([$`])";
        private const string escapeReplaceString = "`$1";
        private const string doubleQuote = "\"";
        private const string escapedDoubleQuote = "`\"";

        public IEnumerable<CompletionResult> CompleteArgument(
            string commandName,
            string parameterName,
            string wordToComplete,
            CommandAst commandAst,
            IDictionary fakeBoundParameters)
        {
            string matchString = wordToComplete.Trim('"', '\'');
            var requireQuotes = new Regex(requireQuotesPattern);

            foreach (string font in Utils.FontList)
            {
                string escapedFont = Regex.Replace(font, requireEscapingPattern, escapeReplaceString);
                if (string.IsNullOrEmpty(wordToComplete)
                    || font.StartsWith(matchString, StringComparison.OrdinalIgnoreCase))
                {
                    if (requireQuotes.IsMatch(font) || wordToComplete.StartsWith(doubleQuote))
                    {
                        var result = string.Format("\"{0}\"", Regex.Replace(escapedFont, doubleQuote, escapedDoubleQuote));
                        yield return new CompletionResult(
                            completionText: result,
                            listItemText: font,
                            CompletionResultType.ParameterValue,
                            toolTip: font);
                        continue;
                    }

                    yield return new CompletionResult(
                        completionText: escapedFont,
                        listItemText: font,
                        CompletionResultType.ParameterName,
                        toolTip: font);
                }
            }
        }
    }
}
