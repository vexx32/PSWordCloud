using System.Collections;
using System.Collections.Generic;
using System.Management.Automation;
using System.Management.Automation.Language;

namespace PSWordCloud.Completers
{
    public class AngleCompleter : IArgumentCompleter
    {
        public IEnumerable<CompletionResult> CompleteArgument(
            string commandName,
            string parameterName,
            string wordToComplete,
            CommandAst commandAst,
            IDictionary fakeBoundParameters)
        {
            for (float angle = 0; angle <= 360; angle += 45)
            {
                var s = angle.ConvertTo<string>();
                yield return new CompletionResult(s, s, CompletionResultType.ParameterValue, s);
            }
        }
    }
}
