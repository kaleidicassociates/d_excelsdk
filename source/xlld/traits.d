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
          procedure: name,
          typeText: typeText,
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
        ret.procedure = ret.functionText = __traits(identifier, F);
        ret.typeText = getTypeText!F;
        ret.macroType = "1"w;
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

struct DllDefFile {
    Statement[] statements;
}

struct Statement {
    string name;
    string[] args;

    this(string name, string[] args) @safe pure nothrow {
        this.name = name;
        this.args = args;
    }

    this(string name, string arg) @safe pure nothrow {
        this(name, [arg]);
    }

    string toString() @safe pure const {
        import std.array: join;
        import std.algorithm: map;

        if(name == "EXPORTS")
            return name ~ "\n" ~ args.map!(a => "\t\t" ~ a).join("\n");
        else
            return name ~ "\t\t" ~ args.map!(a => stringify(name, a)).join(" ");
    }

    static private string stringify(in string name, in string arg) @safe pure {
        if(name == "LIBRARY") return `"` ~ arg ~ `"`;
        if(name == "DESCRIPTION") return `'` ~ arg ~ `'`;
        return arg;
    }
}

/**
   Returns a structure descripting a Windows .def file.
   This allows the tests to not care about the specific formatting
   used when writing the information out.
   This encapsulates all the functions to be exported by the DLL/XLL.
 */
DllDefFile dllDefFile(Modules...)(string libName, string description)
if(allSatisfy!(isSomeString, typeof(Modules)))
{
    import std.conv: to;

    auto statements = [
        Statement("LIBRARY", libName),
        Statement("DESCRIPTION", description),
        Statement("EXETYPE", "NT"),
        Statement("CODE", "PRELOAD DISCARDABLE"),
        Statement("DATA", "PRELOAD MULTIPLE"),
    ];

    string[] exports = ["xlAutoOpen"];
    foreach(func; getAllWorksheetFunctions!Modules) {
        exports ~= func.procedure.to!string;
    }

    return DllDefFile(statements ~ Statement("EXPORTS", exports));
}

@("worksheet functions to .def file")
unittest {
    dllDefFile!"xlld.test_xl_funcs"("myxll32.dll", "Simple D add-in").shouldEqual(
        DllDefFile(
            [
                Statement("LIBRARY", "myxll32.dll"),
                Statement("DESCRIPTION", "Simple D add-in"),
                Statement("EXETYPE", "NT"),
                Statement("CODE", "PRELOAD DISCARDABLE"),
                Statement("DATA", "PRELOAD MULTIPLE"),
                Statement("EXPORTS", ["xlAutoOpen", "FuncMulByTwo", "FuncFP12", "FuncFib"]),
            ]
        )
    );
}


mixin template GenerateDllDef(string module_ = __MODULE__) {
    version(main) {
        void main(string[] args) {

            import std.stdio: File;
            import std.exception: enforce;
            import std.path: stripExtension;

            enforce(args.length >= 2 && args.length <= 4,
                    "Usage: " ~ args[0] ~ " [file_name] <lib_name> <description>");

            immutable fileName = args[1];
            immutable libName = args.length > 2 ? args[2] : fileName.stripExtension ~ ".xll";
            immutable description = args.length > 3 ? args[3] : "Simple D add-in to Excel";

            auto file = File(fileName, "w");
            foreach(stmt; dllDefFile!module_(libName, description).statements)
                file.writeln(stmt.toString);
        }
    }
}
