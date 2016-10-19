module xlld.traits;

import xlld.worksheet;
import std.traits: isSomeFunction, allSatisfy, isSomeString;

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

private template isExcelFunction(alias T) {
    // trying to get a pointer to something is a good way of making sure we can
    // attempt to evaluate `isSomeFunction` - it's not always possible
    enum canGetPointerToIt = __traits(compiles, &T);

    static if(canGetPointerToIt)
        enum isExcelFunction = isSomeFunction!T;
    else
        enum isExcelFunction = false;
}

WorksheetFunction[] getModuleExcelFunctions(string moduleName)() {
    import std.traits: fullyQualifiedName;
    mixin(`import ` ~ fullyQualifiedName!(mixin(moduleName)) ~ `;`);
    alias module_ = Identity!(mixin(moduleName));

    WorksheetFunction[] ret;

    foreach(moduleMemberStr; __traits(allMembers, module_)) {

        alias moduleMember = Identity!(__traits(getMember, module_, moduleMemberStr));

        static if(isExcelFunction!moduleMember)
            ret ~= getWorksheetFunction!moduleMember;
    }

    return ret;
}

@("getExcelFunctions on test_module")
@safe pure unittest {
    getModuleExcelFunctions!"xlld.test_module".shouldEqual(
        [
            doubleToDoubleFunction("FuncMulByTwo"),
        ]
    );
}

WorksheetFunction[] getAllExcelFunctions(Modules...)() if(allSatisfy!(isSomeString, typeof(Modules))) {
    WorksheetFunction[] ret;
    foreach(module_; Modules) {
        ret ~= getModuleExcelFunctions!module_;
    }

    return ret;
}

mixin template implGetWorksheetFunctions(Modules...) if(allSatisfy!(isSomeString, typeof(Modules))) {
    extern(C) WorksheetFunction[] getWorksheetFunctions() @safe pure nothrow {
        return getAllExcelFunctions!Modules;
    }
}

@("template mixin for getWorkSheetFunctions for test_module")
unittest {
    import xlld.traits;
    import xlld.worksheet;

    // mixin the function here then call it to see if it does what it's supposed to
    mixin implGetWorksheetFunctions!"xlld.test_module";
    getWorksheetFunctions.shouldEqual([doubleToDoubleFunction("FuncMulByTwo")]);
}
