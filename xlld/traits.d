module xlld.traits;

import xlld.worksheet;
import std.traits: isSomeFunction;

version(unittest) {
    import unit_threaded;

    WorksheetFunction doubleToDoubleFunction(wstring name) @safe pure nothrow {
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
@safe pure unittest {
    double foo(double) { return 0; }
    getWorksheetFunction!foo.shouldEqual(doubleToDoubleFunction("foo"));

    double bar(double) { return 0; }
    getWorksheetFunction!bar.shouldEqual(doubleToDoubleFunction("bar"));
}

private alias Identity(alias T) = T;


WorksheetFunction[] getExcelFunctions(string moduleName)() {
    import std.traits: fullyQualifiedName;
    mixin(`import ` ~ fullyQualifiedName!(mixin(moduleName)) ~ `;`);
    alias module_ = Identity!(mixin(moduleName));

    WorksheetFunction[] ret;

    foreach(moduleMemberStr; __traits(allMembers, module_)) {
        alias moduleMember = Identity!(__traits(getMember, module_, moduleMemberStr));
        static if(isSomeFunction!moduleMember) {
            ret ~= getWorksheetFunction!moduleMember;
        }
    }

    return ret;
}

@("getExcelFunctions on test_module")
@safe pure unittest {
    getExcelFunctions!"xlld.test_module".shouldEqual(
        [
            doubleToDoubleFunction("FuncMulByTwo"),
        ]
    );
}
