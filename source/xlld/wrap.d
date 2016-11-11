module xlld.wrap;

import xlld.xlcall;
import xlld.traits: isSupportedFunction;

version(unittest) {
    import unit_threaded;
}

XLOPER12 toXlOper(T)(T val) if(is(T == double)) {
    auto ret = XLOPER12();
    ret.xltype = xltypeNum;
    ret.val.num = val;
    return ret;
}

XLOPER12 toXlOper(T)(T val) if(is(T == string)) {
    import std.conv: to;

    // the first wchar is the string length
    auto wval = val.to!wstring;
    wval = [cast(immutable wchar)val.length] ~ wval;

    auto ret = XLOPER12();
    ret.xltype = xltypeStr;
    ret.val.str = cast(XCHAR*)wval.ptr;
    return ret;
}

XLOPER12 toXlOper(T)(T[][] values) if(is(T == double) || is(T == string)) {
    import std.algorithm: map, all;
    import std.array: array;
    import std.exception: enforce;
    import std.conv: text;

    enforce(values.all!(a => a.length == values[0].length),
            text("# of columns must all be the same and aren't: ", values.map!(a => a.length)));

    auto ret = XLOPER12();
    ret.xltype = xltypeMulti;
    ret.val.array.rows = cast(int)values.length;
    ret.val.array.columns = cast(int)values[0].length;

    XLOPER12[] opers;
    foreach(row; values)
        foreach(val; row)
            opers ~= val.toXlOper;

    ret.val.array.lparray = opers.ptr;

    return ret;
}

XLOPER12 toXlOper(T)(T values) if(is(T == string[]) || is(T == double[])) {
    return [values].toXlOper;
}

auto fromXlOper(T)(LPXLOPER12 val) if(is(T == double)) {
    if(val.xltype == xltypeMissing)
        return double.init;

    return val.val.num;
}

// 2D slices
auto fromXlOper(T)(LPXLOPER12 val) if(is(T: E[][], E) && (is(E == string) || is(E == double))) {
    return val.fromXlOperMulti!(typeof(T.init[0][0]));
}

// 1D slices
auto fromXlOper(T)(LPXLOPER12 val) if(is(T: E[], E) && (is(E == string) || is(E == double))) {
    import std.array: join;
    return val.fromXlOperMulti!(typeof(T.init[0])).join;
}

private auto fromXlOperMulti(T)(LPXLOPER12 val) {
    import xlld.xl: coerce, free;
    import std.exception: enforce;
    import std.conv: text;

    enforce(val.xltype == xltypeMulti,
            text("Cannot convert XL oper of type ", val.xltype));

    T[][] ret;

    const rows = val.val.array.rows;
    const columns = val.val.array.columns;

    ret.length = rows;
    foreach(ref col; ret)
        col.length = columns;

    auto values = val.val.array.lparray[0 .. (rows * columns)];

    foreach(const row; 0 .. rows) {
        foreach(const col; 0 .. columns) {
            auto cellVal = coerce(&values[row * columns + col]);
            scope(exit) free(&cellVal);
            if(cellVal.xltype == dlangToXlOperType!T.Type)
                ret[row][col] = (&cellVal).fromXlOper!T;
            else
                ret[row][col] = T.init;
        }
    }

    return ret;
}


auto fromXlOper(T)(LPXLOPER12 val) if(is(T == string)) {
    import std.conv: to;

    if(val.xltype == xltypeMissing)
        return null;

    wchar[] ret;
    ret.length = val.val.str[0];
    ret[0 .. $] = val.val.str[1 .. ret.length + 1];
    return ret.to!string;
}


private enum isWorksheetFunction(alias F) =
    isSupportedFunction!(F, double, double[][], string[][], string[], double[], string);

@safe pure unittest {
    import xlld.test_d_funcs;
    static assert(!isWorksheetFunction!shouldNotBeAProblem);
}

string wrapWorksheetFunctionsString(string moduleName)() {
    if(!__ctfe) {
        return "";
    }

    import xlld.traits: Identity;
    import std.array: join;
    import std.traits: ReturnType, Parameters;

    mixin(`import ` ~ moduleName ~ `;`);
    alias module_ = Identity!(mixin(moduleName));

    string ret = `static import ` ~ moduleName ~ ";\n\n";

    foreach(moduleMemberStr; __traits(allMembers, module_)) {
        alias moduleMember = Identity!(__traits(getMember, module_, moduleMemberStr));

        static if(isWorksheetFunction!moduleMember) {
            ret ~= wrapModuleFunctionStr!(moduleName, moduleMemberStr);
        }
    }

    return ret;
}

version(unittest) {
    // automatically converts from oper to compare with a D type
    void shouldEqualDlang(U)(LPXLOPER12 actual, U expected, string file = __FILE__, size_t line = __LINE__) {
        if(actual.xltype == xltypeErr)
            fail("XLOPER is of error type", file, line);
        actual.fromXlOper!U.shouldEqual(expected, file, line);
    }

    XLOPER12 toSRef(T)(T val) {
        auto ret = toXlOper(val);
        ret.xltype = xltypeSRef;
        return ret;
    }
}

@("Wrap double[][] -> double")
@system unittest {
    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");

    auto arg = toSRef(cast(double[][])[[1, 2, 3, 4], [11, 12, 13, 14]]);
    FuncAddEverything(&arg).shouldEqualDlang(60.0);

    arg = toSRef(cast(double[][])[[0, 1, 2, 3], [4, 5, 6, 7]]);
    FuncAddEverything(&arg).shouldEqualDlang(28.0);
}

@("Wrap double[][] -> double[][]")
@system unittest {
    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");

    auto arg = toSRef(cast(double[][])[[1, 2, 3, 4], [11, 12, 13, 14]]);
    FuncTripleEverything(&arg).shouldEqualDlang(cast(double[][])[[3, 6, 9, 12], [33, 36, 39, 42]]);

    arg = toSRef(cast(double[][])[[0, 1, 2, 3], [4, 5, 6, 7]]);
    FuncTripleEverything(&arg).shouldEqualDlang(cast(double[][])[[0, 3, 6, 9], [12, 15, 18, 21]]);
}


@("Wrap string[][] -> double")
@system unittest {
    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");

    auto arg = toSRef([["foo", "bar", "baz", "quux"], ["toto", "titi", "tutu", "tete"]]);
    FuncAllLengths(&arg).shouldEqualDlang(29.0);

    arg = toSRef([["", "", "", ""], ["", "", "", ""]]);
    FuncAllLengths(&arg).shouldEqualDlang(0.0);
}

@("Wrap string[][] -> double[][]")
@system unittest {
    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");

    auto arg = toSRef([["foo", "bar", "baz", "quux"], ["toto", "titi", "tutu", "tete"]]);
    FuncLengths(&arg).shouldEqualDlang(cast(double[][])[[3, 3, 3, 4], [4, 4, 4, 4]]);

    arg = toSRef([["", "", ""], ["", "", "huh"]]);
    FuncLengths(&arg).shouldEqualDlang(cast(double[][])[[0, 0, 0], [0, 0, 3]]);
}

@("Wrap string[][] -> string[][]")
@system unittest {
    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");

    auto arg = toSRef([["foo", "bar", "baz", "quux"], ["toto", "titi", "tutu", "tete"]]);
    FuncBob(&arg).shouldEqualDlang([["foobob", "barbob", "bazbob", "quuxbob"],
                                    ["totobob", "titibob", "tutubob", "tetebob"]]);
}

@("Wrap string[] -> double")
@system unittest {
    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");
    auto arg = toSRef([["foo", "bar"], ["baz", "quux"]]);
    FuncStringSlice(&arg).shouldEqualDlang(4.0);
}

@("Wrap double[] -> double")
@system unittest {
    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");
    auto arg = toSRef([[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]]);
    FuncDoubleSlice(&arg).shouldEqualDlang(6.0);
}

@("Wrap double[] -> double[]")
@system unittest {
    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");
    auto arg = toSRef([[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]]);
    FuncSliceTimes3(&arg).shouldEqualDlang([3.0, 6.0, 9.0, 12.0, 15.0, 18.0]);
}

@("Wrap string[] -> string[]")
@system unittest {
    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");
    auto arg = toSRef(["quux", "toto"]);
    StringsToStrings(&arg).shouldEqualDlang(["quuxfoo", "totofoo"]);
}

@("Wrap string[] -> string")
@system unittest {
    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");
    auto arg = toSRef(["quux", "toto"]);
    StringsToString(&arg).shouldEqualDlang("quux, toto");
}

@("Wrap string -> string")
@system unittest {
    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");
    auto arg = toXlOper("foo");
    StringToString(&arg).shouldEqualDlang("foobar");
}

@("Wrap string, string, string -> string")
@system unittest {
    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");
    auto arg0 = toXlOper("foo");
    auto arg1 = toXlOper("bar");
    auto arg2 = toXlOper("baz");
    ManyToString(&arg0, &arg1, &arg2).shouldEqualDlang("foobarbaz");
}

private enum invalidXlOperType = 0xdeadbeef;

/**
 Maps a D type to two integer xltypes from XLOPER12.
 InputType is the type actually passed in by the spreadsheet,
 whilst Type is the Type that it gets coerced to.
 */
template dlangToXlOperType(T) {
    static if(is(T == double[][]) || is(T == string[][]) || is(T == double[]) || is(T == string[])) {
        enum InputType = xltypeSRef;
        enum Type= xltypeMulti;
    } else static if(is(T == double)) {
        enum InputType = xltypeNum;
        enum Type = xltypeNum;
    } else static if(is(T == string)) {
        enum InputType = xltypeStr;
        enum Type = xltypeStr;
    } else {
        enum InputType = invalidXlOperType;
        enum Type = invalidXlOperType;
    }
}

string wrapModuleFunctionStr(string moduleName, string funcName)() {
    if(!__ctfe) {
        return "";
    }

    import std.array: join;
    import std.traits: Parameters;
    import std.conv: to;
    import std.algorithm: map;
    import std.range: iota;
    mixin("import " ~ moduleName ~ ": " ~ funcName ~ ";");

    const argsLength = Parameters!(mixin(funcName)).length;
    // e.g. LPXLOPER12 arg0, LPXLOPER12 arg1, ...
    const argsDecl = argsLength.iota.map!(a => `LPXLOPER12 arg` ~ a.to!string).join(", ");
    // e.g. arg0, arg1, ...
    const argsCall = argsLength.iota.map!(a => `arg` ~ a.to!string).join(", ");

    return [
        `extern(Windows) LPXLOPER12 ` ~ funcName ~ `(` ~ argsDecl ~ `) {`,
        `    static import ` ~ moduleName ~ `;`,
        `    alias wrappedFunc = ` ~ moduleName ~ `.` ~ funcName ~ `;`,
        `    return wrapModuleFunctionImpl!wrappedFunc(` ~ argsCall ~  `);`,
        `}`,
    ].join("\n");
}

LPXLOPER12 wrapModuleFunctionImpl(alias wrappedFunc, T...)(T args) {
    import xlld.xl: free;
    import std.conv: text;
    import std.traits: Parameters;
    import std.typecons: Tuple;

    static XLOPER12 ret;

    XLOPER12[T.length] realArgs;
    // must 1st convert each argument to the "real" type.
    // 2D arrays are passed in as SRefs, for instance
    foreach(i, InputType; Parameters!wrappedFunc) {
        if(args[i].xltype == xltypeMissing) {
             realArgs[i] = *args[i];
             continue;
        }
        try
            realArgs[i] = convertInput!InputType(args[i]);
        catch(Exception ex) {
            ret.xltype = xltypeErr;
            ret.val.err = -1;
            return &ret;
        }
    }

    scope(exit)
        foreach(ref arg; realArgs)
            free(&arg);

    Tuple!(Parameters!wrappedFunc) dArgs; // the D types to pass to the wrapped function

    // next call the wrapped function with D types
    foreach(i, InputType; Parameters!wrappedFunc) {
        try
            dArgs[i] = fromXlOper!InputType(&realArgs[i]);
        catch(Exception ex) {
            ret.xltype = xltypeErr;
            ret.val.err = -1;
            return &ret;
        }
    }

    ret = toXlOper(wrappedFunc(dArgs.expand));
    return &ret;
}


string wrapAll(string OriginalModule = __MODULE__, Modules...)() {

    if(!__ctfe) {
        return "";
    }

    import xlld.traits: implGetWorksheetFunctionsString;
    return
        wrapWorksheetFunctionsString!Modules ~
        "\n" ~
        implGetWorksheetFunctionsString!OriginalModule ~
        "\n" ~
        `mixin GenerateDllDef!"` ~ OriginalModule ~ `";` ~
        "\n";
}


XLOPER12 convertInput(T)(LPXLOPER12 arg) {
    import xlld.xl: coerce, free;

    if(arg.xltype != dlangToXlOperType!T.InputType)
        throw new Exception("Wrong input type");

    auto realArg = coerce(arg);

    if(realArg.xltype != dlangToXlOperType!T.Type) {
        free(&realArg);
        throw new Exception("Wrong converted input type");
    }

    return realArg;
}
