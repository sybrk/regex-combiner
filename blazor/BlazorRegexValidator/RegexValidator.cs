using System.Text.RegularExpressions;
using Microsoft.JSInterop;

public static class RegexValidator
{
    [JSInvokable("ValidateRegex")]
    public static bool MatchRegex(string pattern)
    {
        try
        {
            var match = Regex.Match("Hello world", pattern);
            return true;
        }
        catch (System.Exception)
        {

            return false;
        }
        
        
    }
}
