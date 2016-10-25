/**
   Interface for registering worksheet functions with Excel
 */
module xlld.worksheet;

/**
 Simple wrapper struct for a value. Provides a type-safe way
 of making sure positional arguments match the intended semantics,
 which is important given that nearly all of the arguments for
 worksheet function registration are of the same type: wstring
 */
mixin template SmallType(string name, T = wstring) {
    mixin(`struct ` ~ name ~ `{ T value; }`);
}

mixin SmallType!"Procedure";
mixin SmallType!"TypeText";
mixin SmallType!"FunctionText";
mixin SmallType!"ArgumentText";
mixin SmallType!"MacroType";
mixin SmallType!"Category";
mixin SmallType!"ShortcutText";
mixin SmallType!"HelpTopic";
mixin SmallType!"FunctionHelp";
mixin SmallType!("ArgumentHelp", wstring[]);


struct WorksheetFunction {
    Procedure procedure;
    TypeText typeText;
    FunctionText functionText;
    Optional optional;
    alias optional this;

    const(wstring)[] toStringArray() @safe pure const nothrow {
        return [procedure.value, typeText.value,
                functionText.value, argumentText.value,
                macroType.value, category.value,
                shortcutText.value, helpTopic.value, functionHelp.value] ~ argumentHelp.value;
    }
}

struct Optional {
    ArgumentText argumentText;
    MacroType macroType = MacroType("1"w);
    Category category;
    ShortcutText shortcutText;
    HelpTopic helpTopic;
    FunctionHelp functionHelp;
    ArgumentHelp argumentHelp;
}

alias Register = Optional;
