/**
 This module implements the compile-time reflection machinery to
 automatically register all D functions that are eligible in a
 compile-time define list of modules to be called from Excel.

 Import this module from any module from your XLL build and:

 -----------
 import xlld;

 mixin(implGetWorksheetFunctionsString!("module1", "module2", "module3"));
 -----------

 All eligible functions in the 3 example modules above will automagically
 be accessible from Excel (assuming the built XLL is loaded as an add-in).
 */
module xlld.traits;

import xlld.worksheet;
import std.traits: isSomeFunction, allSatisfy, isSomeString;

// import unit_threaded and introduce helper functions for testing
version(unittest) {
    import unit_threaded;

    // return a WorksheetFunction for a double function(double) with no
    // optional arguments
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

/**
 Take a D function as a compile-time parameter and returns a
 WorksheetFunction struct with the fields filled in accordingly.
 */
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

// helper template for aliasing
private alias Identity(alias T) = T;

// whether or not this is a function that can be called from Excel
private template isWorksheetFunction(alias T) {
    // trying to get a pointer to something is a good way of making sure we can
    // attempt to evaluate `isSomeFunction` - it's not always possible
    enum canGetPointerToIt = __traits(compiles, &T);

    static if(canGetPointerToIt)
        enum isWorksheetFunction = isSomeFunction!T;
    else
        enum isWorksheetFunction = false;
}

/**
 Gets all Excel-callable functions in a given module
 */
WorksheetFunction[] getModuleWorksheetFunctions(string moduleName)() {
    mixin(`import ` ~ moduleName ~ `;`);
    alias module_ = Identity!(mixin(moduleName));

    WorksheetFunction[] ret;

    foreach(moduleMemberStr; __traits(allMembers, module_)) {

        alias moduleMember = Identity!(__traits(getMember, module_, moduleMemberStr));

        static if(isWorksheetFunction!moduleMember)
            ret ~= getWorksheetFunction!moduleMember;
    }

    return ret;
}

@("getWorksheetFunctions on test_module")
@safe pure unittest {
    getModuleWorksheetFunctions!"xlld.test_module".shouldEqual(
        [
            doubleToDoubleFunction("FuncMulByTwo"),
        ]
    );
}

/**
 Gets all Excel-callable functions from the given modules
 */
WorksheetFunction[] getAllWorksheetFunctions(Modules...)() if(allSatisfy!(isSomeString, typeof(Modules))) {
    WorksheetFunction[] ret;

    foreach(module_; Modules) {
        ret ~= getModuleWorksheetFunctions!module_;
    }

    return ret;
}

/**
 Implements the getWorksheetFunctions function needed by xlld.xll in
 order to register the Excel-callable functions at runtime
 This used to be a template mixin but even using a string mixin inside
 fails to actually make it an extern(C) function.
 */
string implGetWorksheetFunctionsString(Modules...)() if(allSatisfy!(isSomeString, typeof(Modules))) {
    import std.array: join;

    string modulesString() {

        string[] modules;
        foreach(module_; Modules) {
            modules ~= `"` ~ module_ ~ `"`;
        }
        return modules.join(", ");
    }

    return
        [
            `extern(C) WorksheetFunction[] getWorksheetFunctions() @safe pure nothrow {`,
            `    return getAllWorksheetFunctions!(` ~ modulesString ~ `);`,
            `}`,
        ].join("\n");
}

@("template mixin for getWorkSheetFunctions for test_module")
unittest {
    import xlld.traits;
    import xlld.worksheet;

    // mixin the function here then call it to see if it does what it's supposed to
    mixin(implGetWorksheetFunctionsString!"xlld.test_module");
    getWorksheetFunctions.shouldEqual([doubleToDoubleFunction("FuncMulByTwo")]);
}
