module xlld.traits;

import xlld.worksheet;
import std.traits: isSomeFunction;

version(unittest) {
    import unit_threaded;

    WorksheetFunction doubleToDoubleFunction(wstring name) {
        WorksheetFunction func = {
          procedure: name,
          typeText: "BB"w,
          functionText: name,
          argumentText: ""w,
          macroType: "1"w,
          category: ""w,
          shortcutText: ""w,
          helpTopic: ""w,
          functionHelp: ""w,
          argumentHelp: [],
        };

        return func;
    }
}

WorksheetFunction getWorksheetFunction(alias F)() if(isSomeFunction!F) {
    WorksheetFunction ret;
    ret.procedure = ret.functionText = __traits(identifier, F);
    ret.typeText = "BB"w;
    ret.macroType = "1"w;
    return ret;
}

@("double -> double functions with no extra attributes")
unittest {
    double foo(double) { return 0; }
    getWorksheetFunction!foo.shouldEqual(doubleToDoubleFunction("foo"));

    double bar(double) { return 0; }
    getWorksheetFunction!bar.shouldEqual(doubleToDoubleFunction("bar"));
}
