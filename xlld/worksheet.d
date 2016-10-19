/**
   Interface for registering worksheet functions with Excel
 */
module xlld.worksheet;

struct WorksheetFunction {
    wstring procedure;
    wstring typeText;
    wstring functionText;
    wstring argumentText;
    wstring macroType;
    wstring category;
    wstring shortcutText;
    wstring helpTopic;
    wstring functionHelp;
    wstring[] argumentHelp;

    const(wstring)[] toStringArray() @safe pure const nothrow {
        return [procedure, typeText, functionText, argumentText,
                macroType, category, shortcutText, helpTopic,
                functionHelp] ~ argumentHelp;
    }
}
