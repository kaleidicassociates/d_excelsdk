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

private auto fromXlOperMulti(T, int XlType)(LPXLOPER12 val) {
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
            auto cellVal = values[row * columns + col];
            enforce(cellVal.xltype == XlType,
                    text("Unsupported element type in multi: ", cellVal.xltype));
            ret[row][col] = (&cellVal).fromXlOper!T;
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


private enum isWorksheetFunction(alias F) = isSupportedFunction!(F, double, double[][], string[][]);

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
            immutable prmTypeStr = Parameters!moduleMember.stringof;
            ret ~=
                [
                    `extern(Windows) LPXLOPER12 ` ~ moduleMemberStr ~ `(LPXLOPER12 arg) {`,
                    `    static import ` ~ moduleName ~ `;`,
                    `    static XLOPER12 ret;`,
                    `    ret = ` ~ moduleName ~ `.` ~ moduleMemberStr ~
                    `(arg.fromXlOper!` ~ prmTypeStr ~ `).toXlOper;`,
                    `    return &ret;`,
                    `}`,
                    ``,
                ].join("\n");
        }
    }

    return ret;
}

version(unittest) {
    // automatically converts from oper to compare with a D type
    void shouldEqualDlang(T, U)(T actual, U expected, string file = __FILE__, ulong line = __LINE__) {
        actual.fromXlOper!U.shouldEqual(expected);
    }
}

@("Wrap double[][] -> double")
@system unittest {
    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");

    auto arg = toXlOper(cast(double[][])[[1, 2, 3, 4], [11, 12, 13, 14]]);
    FuncAddEverything(&arg).shouldEqualDlang(60.0);

    arg = toXlOper(cast(double[][])[[0, 1, 2, 3], [4, 5, 6, 7]]);
    FuncAddEverything(&arg).shouldEqualDlang(28.0);
}

@("Wrap double[][] -> double[][]")
@system unittest {
    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");

    auto arg = toXlOper(cast(double[][])[[1, 2, 3, 4], [11, 12, 13, 14]]);
    FuncTripleEverything(&arg).shouldEqualDlang(cast(double[][])[[3, 6, 9, 12], [33, 36, 39, 42]]);

    arg = toXlOper(cast(double[][])[[0, 1, 2, 3], [4, 5, 6, 7]]);
    FuncTripleEverything(&arg).shouldEqualDlang(cast(double[][])[[0, 3, 6, 9], [12, 15, 18, 21]]);
}


@("Wrap string[][] -> double")
@system unittest {
    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");

    auto arg = toXlOper([["foo", "bar", "baz", "quux"], ["toto", "titi", "tutu", "tete"]]);
    FuncAllLengths(&arg).shouldEqualDlang(29.0);

    arg = toXlOper([["", "", "", ""], ["", "", "", ""]]);
    FuncAllLengths(&arg).shouldEqualDlang(0.0);
}

@("Wrap string[][] -> double[][]")
@system unittest {
    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");

    auto arg = toXlOper([["foo", "bar", "baz", "quux"], ["toto", "titi", "tutu", "tete"]]);
    FuncLengths(&arg).shouldEqualDlang(cast(double[][])[[3, 3, 3, 4], [4, 4, 4, 4]]);

    arg = toXlOper([["", "", ""], ["", "", "huh"]]);
    FuncLengths(&arg).shouldEqualDlang(cast(double[][])[[0, 0, 0], [0, 0, 3]]);
}

@("Wrap string[][] -> string[][]")
@system unittest {
    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");

    auto arg = toXlOper([["foo", "bar", "baz", "quux"], ["toto", "titi", "tutu", "tete"]]);
    FuncBob(&arg).shouldEqualDlang([["foobob", "barbob", "bazbob", "quuxbob"],
                                    ["totobob", "titibob", "tutubob", "tetebob"]]);
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
    else static if(is(T == double))
        enum dlangToXlOperType = xltypeNum;
    else
        enum dlangToXlOperType = invalidXlOperType;
}
