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
import xlld.xlcall;
import std.traits: isSomeFunction, allSatisfy, isSomeString;

// import unit_threaded and introduce helper functions for testing
version(unittest) {
    import unit_threaded;

    // return a WorksheetFunction for a double function(double) with no
    // optional arguments
    WorksheetFunction makeWorksheetFunction(wstring name, wstring typeText) @safe pure nothrow {
        WorksheetFunction func = {
          procedure: Procedure(name),
          typeText: TypeText(typeText),
          functionText: FunctionText(name),
          argumentText: ArgumentText(""w),
          macroType: MacroType("1"w),
          category: Category(""w),
          shortcutText: ShortcutText(""w),
          helpTopic: HelpTopic(""w),
          functionHelp: FunctionHelp(""w),
          argumentHelp: ArgumentHelp([]),
        };

        return func;
    }

    WorksheetFunction doubleToDoubleFunction(wstring name) @safe pure nothrow {
        return makeWorksheetFunction(name, "BB"w);
    }

    WorksheetFunction FP12ToDoubleFunction(wstring name) @safe pure nothrow {
        return makeWorksheetFunction(name, "BK%"w);
    }

    WorksheetFunction operToOperFunction(wstring name) @safe pure nothrow {
        return makeWorksheetFunction(name, "UU"w);
    }
}

/**
 Take a D function as a compile-time parameter and returns a
 WorksheetFunction struct with the fields filled in accordingly.
 */
WorksheetFunction getWorksheetFunction(alias F)() if(isSomeFunction!F) {
    import std.traits: ReturnType, Parameters;

    alias R = ReturnType!F;
    alias T = Parameters!F;

    static if(!isWorksheetFunction!F) {
        throw new Exception("Unsupported function type " ~ R.stringof ~ T.stringof ~ " for " ~
                            __traits(identifier, F).stringof[1 .. $-1]);
    } else {

        WorksheetFunction ret;
        ret.procedure = Procedure(__traits(identifier, F));
        ret.functionText = FunctionText(__traits(identifier, F));
        ret.typeText = TypeText(getTypeText!F);
        ret.macroType = MacroType("1"w);
        return ret;
    }
}

@("getWorksheetFunction for double -> double functions with no extra attributes")
@safe pure unittest {
    double foo(double) { return 0; }
    getWorksheetFunction!foo.shouldEqual(doubleToDoubleFunction("foo"));

    double bar(double) { return 0; }
    getWorksheetFunction!bar.shouldEqual(doubleToDoubleFunction("bar"));
}

@("getWorksheetFunction for double -> int functions should fail")
@safe pure unittest {
    double foo(int) { return 0; }
    getWorksheetFunction!foo.shouldThrowWithMessage("Unsupported function type double(int) for foo");
}


private wstring getTypeText(alias F)() if(isSomeFunction!F) {
    import std.traits: ReturnType, Parameters;

    wstring typeToString(T)() {
        static if(is(T == double))
            return "B";
        else static if(is(T == FP12*))
            return "K%";
        else static if(is(T == LPXLOPER12))
            return "U";
        else
            static assert(false, "Unsupported type " ~ T.stringof);
    }

    wstring retType = typeToString!(ReturnType!F);
    foreach(argType; Parameters!F)
        retType ~= typeToString!(argType);

    return retType;
}


@("getTypeText")
@safe pure unittest {
    import std.conv: to; // working around unit-threaded bug

    double foo(double);
    getTypeText!foo.to!string.shouldEqual("BB");

    double bar(FP12*);
    getTypeText!bar.to!string.shouldEqual("BK%");

    FP12* baz(FP12*);
    getTypeText!baz.to!string.shouldEqual("K%K%");

    FP12* qux(double);
    getTypeText!qux.to!string.shouldEqual("K%B");

    LPXLOPER12 fun(LPXLOPER12);
    getTypeText!fun.to!string.shouldEqual("UU");
}



// helper template for aliasing
private alias Identity(alias T) = T;


// whether or not this is a function that has the "right" types
template isSupportedFunction(alias F, T...) {
    import std.traits: isSomeFunction, ReturnType, Parameters;
    import std.meta: AliasSeq, allSatisfy;

    // trying to get a pointer to something is a good way of making sure we can
    // attempt to evaluate `isSomeFunction` - it's not always possible
    enum canGetPointerToIt = __traits(compiles, &F);
    enum isOneOfSupported(U) = isSupportedType!(U, T);

    static if(canGetPointerToIt)
        enum isSupportedFunction =
            isSomeFunction!F &&
            isOneOfSupported!(ReturnType!F) &&
            allSatisfy!(isOneOfSupported, Parameters!F);
    else
        enum isSupportedFunction = false;
}

private template isSupportedType(T, U...) {
    static if(U.length == 0)
        enum isSupportedType = false;
    else
        enum isSupportedType = is(T == U[0]) || isSupportedType!(T, U[1..$]);
}

static assert(isSupportedType!(int, int, int));
static assert(!isSupportedType!(int, double, string));


// whether or not this is a function that can be called from Excel
private enum isWorksheetFunction(alias F) = isSupportedFunction!(F, double, FP12*, LPXLOPER12);

/**
 Gets all Excel-callable functions in a given module
 */
WorksheetFunction[] getModuleWorksheetFunctions(string moduleName)() {
    mixin(`import ` ~ moduleName ~ `;`);
    alias module_ = Identity!(mixin(moduleName));

    WorksheetFunction[] ret;

    foreach(moduleMemberStr; __traits(allMembers, module_)) {

        alias moduleMember = Identity!(__traits(getMember, module_, moduleMemberStr));

        static if(isWorksheetFunction!moduleMember) {
            try
                ret ~= getWorksheetFunction!moduleMember;
            catch(Exception ex)
                assert(0); //can't happen
        }
    }

    return ret;
}

@("getWorksheetFunctions on test_xl_funcs")
@safe pure unittest {
    getModuleWorksheetFunctions!"xlld.test_xl_funcs".shouldEqual(
        [
            doubleToDoubleFunction("FuncMulByTwo"),
            FP12ToDoubleFunction("FuncFP12"),
            operToOperFunction("FuncFib"),
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

@("template mixin for getWorkSheetFunctions for test_xl_funcs")
unittest {
    import xlld.traits;
    import xlld.worksheet;

    // mixin the function here then call it to see if it does what it's supposed to
    mixin(implGetWorksheetFunctionsString!"xlld.test_xl_funcs");
    getWorksheetFunctions.shouldEqual(
        [
            doubleToDoubleFunction("FuncMulByTwo"),
            FP12ToDoubleFunction("FuncFP12"),
            operToOperFunction("FuncFib"),
        ]
    );
}
