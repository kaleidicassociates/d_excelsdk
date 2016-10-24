module xlld.newwrap;

import xlld.xlcall;

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


auto fromXlOper(T)(LPXLOPER12 val) if(is(T == double[][])) {
    import std.exception: enforce;
    import std.conv: text;

    enforce(val.xltype == xltypeMulti,
            text("Cannot convert XL oper of type ", val.xltype));

    double[][] ret;

    const rows = val.val.array.rows;
    const columns = val.val.array.columns;

    ret.length = rows;
    foreach(ref col; ret)
        col.length = columns;

    auto values = val.val.array.lparray[0 .. (rows * columns)];

    foreach(const row; 0 .. rows) {
        foreach(const col; 0 .. columns) {
            auto cellVal = values[row * columns + col];
            enforce(cellVal.xltype == xltypeNum,
                    text("Unsupported element type in multi: ", cellVal.xltype));
            ret[row][col] = cellVal.val.num;
        }
    }

    return ret;
}

auto fromXlOper(T)(LPXLOPER12 val) if(is(T == string[][])) {
    import std.exception: enforce;
    import std.conv: text;

    enforce(val.xltype == xltypeMulti,
            text("Cannot convert XL oper of type ", val.xltype));

    string[][] ret;

    const rows = val.val.array.rows;
    const columns = val.val.array.columns;

    ret.length = rows;
    foreach(ref col; ret)
        col.length = columns;

    auto values = val.val.array.lparray[0 .. (rows * columns)];

    foreach(const row; 0 .. rows) {
        foreach(const col; 0 .. columns) {
            auto cellVal = values[row * columns + col];
            enforce(cellVal.xltype == xltypeStr,
                    text("Unsupported element type in multi: ", cellVal.xltype));
            ret[row][col] = (&cellVal).fromXlOper!string;
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


// helper template for aliasing
private alias Identity(alias T) = T;

private template isWorksheetFunction(alias T) {
    import std.traits: isSomeFunction, ReturnType, Parameters;
    import std.meta: AliasSeq, allSatisfy;

    // trying to get a pointer to something is a good way of making sure we can
    // attempt to evaluate `isSomeFunction` - it's not always possible
    enum canGetPointerToIt = __traits(compiles, &T);
    enum isSupportedType(U) = is(U == double) || is(U == double[][]) || is(U == string[][]);

    static if(canGetPointerToIt)
        enum isWorksheetFunction =
            isSomeFunction!T &&
            isSupportedType!(ReturnType!T) &&
            allSatisfy!(isSupportedType, Parameters!T);
    else
        enum isWorksheetFunction = false;

}

string wrapWorksheetFunctionsString(string moduleName)() {
    import std.array: join;
    import std.traits: Parameters;

    mixin(`import ` ~ moduleName ~ `;`);
    alias module_ = Identity!(mixin(moduleName));

    string ret;

    foreach(moduleMemberStr; __traits(allMembers, module_)) {
        alias moduleMember = Identity!(__traits(getMember, module_, moduleMemberStr));

        static if(isWorksheetFunction!moduleMember) {

            ret ~=
                [
                    `static import ` ~ moduleName ~ `;`,
                    `extern(Windows) double ` ~ moduleMemberStr ~ `(LPXLOPER12 arg) {`,
                    `    return ` ~ moduleName ~ `.` ~ moduleMemberStr ~ `(arg.fromXlOper!` ~ Parameters!moduleMember.stringof ~ `);`
                    `}`,
                    ``,
                ].join("\n");
        }
    }

    return ret;
}

@("Wrap double[][] -> double")
@system unittest {
     mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");

    auto arg = toXlOper(cast(double[][])[[1, 2, 3, 4], [11, 12, 13, 14]]);
    FuncAddEverything(&arg).shouldEqual(60);

    arg = toXlOper(cast(double[][])[[0, 1, 2, 3], [4, 5, 6, 7]]);
    FuncAddEverything(&arg).shouldEqual(28);
}


@("Wrap string[][] -> double")
@system unittest {
    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");

    auto arg = toXlOper([["foo", "bar", "baz", "quux"], ["toto", "titi", "tutu", "tete"]]);
    FuncAllLengths(&arg).shouldEqual(29);

    arg = toXlOper([["", "", "", ""], ["", "", "", ""]]);
    FuncAllLengths(&arg).shouldEqual(0);
}
