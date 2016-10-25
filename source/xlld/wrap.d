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

XLOPER12 toXlOper(T)(T[][] values) {
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

auto fromXlOper(T)(LPXLOPER12 val) if(is(T == double)) {
    return val.val.num;
}

auto fromXlOper(T)(LPXLOPER12 val) if(is(T == double[][])) {
    return val.fromXlOperMulti!(double, xltypeNum);
}

auto fromXlOper(T)(LPXLOPER12 val) if(is(T == string[][])) {
    return val.fromXlOperMulti!(string, xltypeStr);
}

auto fromXlOper(T)(LPXLOPER12 val) if(is(T == string[])) {
    import std.array: join;
    return val.fromXlOperMulti!(string, xltypeStr).join;
}

private auto fromXlOperMulti(T, int XlType)(LPXLOPER12 val) {
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
            if(cellVal.xltype == XlType)
                ret[row][col] = (&cellVal).fromXlOper!T;
            else
                ret[row][col] = T.init;
        }
    }

    return ret;
}


auto fromXlOper(T)(LPXLOPER12 val) if(is(T == string)) {
    import std.conv: to;
    wchar[] ret;
    ret.length = val.val.str[0];
    ret[0 .. $] = val.val.str[1 .. ret.length + 1];
    return ret.to!string;
}


private enum isWorksheetFunction(alias F) =
    isSupportedFunction!(F, double, double[][], string[][], string[], double[]);

string wrapWorksheetFunctionsString(string moduleName)() {
    import xlld.traits: Identity;
    import std.array: join;
    import std.traits: ReturnType, Parameters;

    mixin(`import ` ~ moduleName ~ `;`);
    alias module_ = Identity!(mixin(moduleName));

    string ret = `static import ` ~ moduleName ~ ";\n\n";

    foreach(moduleMemberStr; __traits(allMembers, module_)) {
        alias moduleMember = Identity!(__traits(getMember, module_, moduleMemberStr));

        static if(isWorksheetFunction!moduleMember) {
            ret ~= wrapModuleFunctionStr(moduleName, moduleMemberStr);
        }
    }

    return ret;
}

version(unittest) {
    // automatically converts from oper to compare with a D type
    void shouldEqualDlang(T, U)(T actual, U expected, string file = __FILE__, ulong line = __LINE__) {
        actual.xltype.shouldNotEqual(xltypeErr);
        actual.fromXlOper!U.shouldEqual(expected);
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

@("Wrap string[] -> double]")
@system unittest {
    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");
    auto arg = toSRef([["foo", "bar"], ["baz", "quux"]]);
    FuncStringSlice(&arg).shouldEqualDlang(4.0);
}

private enum invalidXlOperType = 0xdeadbeef;

/**
 Maps a D type to an integer xltype in XLOPER12 that would
 get passed in by Excel. This template exists because the values
 that get passed in don't immediately correspond to what would
 be expected. For instance, an xltypeMulti gets passed in as
 an xltypeSRef that must be _coerced_ to an xltypeMulti
 */
template dlangToXlOperInputType(T) {
    static if(is(T == double[][]))
        enum dlangToXlOperInputType = xltypeSRef;
    else static if(is(T == string[][]))
        enum dlangToXlOperInputType = xltypeSRef;
    else static if(is(T == double[]))
        enum dlangToXlOperInputType = xltypeSRef;
    else static if(is(T == string[]))
        enum dlangToXlOperInputType = xltypeSRef;
    else
        enum dlangToXlOperInputType = invalidXlOperType;
}

/**
 Maps a D type to an integer xltype in XLOPER12 that would
 get coerced to after Excel passes it in as input
 */
template dlangToXlOperType(T) {
    static if(is(T == double[][]))
        enum dlangToXlOperType = xltypeMulti;
    else static if(is(T == string[][]))
        enum dlangToXlOperType = xltypeMulti;
    else static if(is(T == double[]))
        enum dlangToXlOperType = xltypeMulti;
    else static if(is(T == string[]))
        enum dlangToXlOperType = xltypeMulti;
    else static if(is(T == double))
        enum dlangToXlOperType = xltypeNum;
    else
        enum dlangToXlOperType = invalidXlOperType;
}


string wrapModuleFunctionStr(string moduleName, string funcName) {
    import std.array: join;
    return [
        `extern(Windows) LPXLOPER12 ` ~ funcName ~ `(LPXLOPER12 arg) {`,
        `    static import ` ~ moduleName ~ `;`,
        `    alias wrappedFunc = ` ~ moduleName ~ `.` ~ funcName ~ `;`,
        `    return wrapModuleFunctionImpl!wrappedFunc(arg);`,
        `}`,
    ].join("\n");
}

LPXLOPER12 wrapModuleFunctionImpl(alias wrappedFunc)(LPXLOPER12 arg) {
    import std.conv: text;
    import std.traits: Parameters;
    import xlld.xl: free, convertInput;

    static assert(Parameters!wrappedFunc.length == 1,
                  text("Illegal number of parameters, only 1 supported, not ",
                       Parameters!wrappedFunc.length));
    alias InputType = Parameters!wrappedFunc[0];
    static XLOPER12 ret;
    // must 1st convert argument to the "real" type.`,
    // 2D arrays are passed in as SRefs, for instance`,
    XLOPER12 realArg;
    try {
        realArg = convertInput!InputType(arg);
    } catch(Exception ex) {
        ret.xltype = xltypeErr;
        ret.val.err = -1;
        return &ret;
    }
    scope(exit) free(&realArg);
    ret = toXlOper(wrappedFunc(fromXlOper!InputType(&realArg)));
    return &ret;
}


string wrapAll(string OriginalModule = __MODULE__, Modules...)() {
    import xlld.traits: implGetWorksheetFunctionsString;
    return
        wrapWorksheetFunctionsString!Modules ~
        "\n" ~
        implGetWorksheetFunctionsString!OriginalModule;
}
