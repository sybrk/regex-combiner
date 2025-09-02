using System.Text.RegularExpressions;
using Microsoft.JSInterop;

public class RegexResult
{
    public bool Success { get; set; }
    public string? Error { get; set; }
}
public static class RegexValidator
{
    [JSInvokable("ValidateRegexDetailed")]
    public static RegexResult MatchRegexDetailed(string pattern)
    {
        try
        {
            _ = Regex.Match("Hello world", pattern);
            return new RegexResult { Success = true, Error = null };
        }
        catch (System.Exception ex)
        {
            return new RegexResult { Success = false, Error = ex.Message };
        }


    }

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
