module xlld.wrap;

import xlld.xlcall;
import xlld.traits: isSupportedFunction;
static import xlld.memorymanager;
import xlld.framework: FreeXLOper;


version(unittest) {
    import unit_threaded;

    // automatically converts from oper to compare with a D type
    void shouldEqualDlang(U)(LPXLOPER12 actual, U expected, string file = __FILE__, size_t line = __LINE__) {
        if(actual.xltype == xltypeErr)
            fail("XLOPER is of error type", file, line);
        actual.fromXlOper!U.shouldEqual(expected, file, line);
    }

    // automatically converts from oper to compare with a D type
    void shouldEqualDlang(U)(ref XLOPER12 actual, U expected, string file = __FILE__, size_t line = __LINE__) {
        shouldEqualDlang(&actual, expected, file, line);
    }

    XLOPER12 toSRef(T)(T val) {
        auto ret = toXlOper(val);
        ret.xltype = xltypeSRef;
        return ret;
    }

    // tracks allocations and throws in the destructor if there is a memory leak
    // it also throws when there is an attempt to deallocate memory that wasn't
    // allocated
    struct TestAllocator {
        import std.experimental.allocator.common: platformAlignment;
        import std.experimental.allocator.mallocator: Mallocator;

        alias allocator = Mallocator.instance;

        private static struct ByteRange {
            void* ptr;
            size_t length;
        }
        private ByteRange[] _allocations;
        private int _numAllocations;

        enum uint alignment = platformAlignment;

        void[] allocate(size_t numBytes) {
            ++_numAllocations;
            auto ret = allocator.allocate(numBytes);
            writelnUt("+ Allocated  ptr ", ret.ptr, " of ", ret.length, " bytes length");
            _allocations ~= ByteRange(ret.ptr, ret.length);
            return ret;
        }

        bool deallocate(void[] bytes) {
            import std.algorithm: remove, canFind;
            import std.exception: enforce;
            import std.conv: text;

            writelnUt("- Deallocate ptr ", bytes.ptr, " of ", bytes.length, " bytes length");

            bool pred(ByteRange other) { return other.ptr == bytes.ptr && other.length == bytes.length; }

            enforce(_allocations.canFind!pred,
                    text("Unknown deallocate byte range. Ptr: ", bytes.ptr, " length: ", bytes.length,
                         " allocations: ", _allocations));
            _allocations = _allocations.remove!pred;
            return allocator.deallocate(bytes);
        }

        auto numAllocations() @safe pure nothrow const {
            return _numAllocations;
        }

        ~this() {
            import std.exception: enforce;
            import std.conv: text;
            enforce(!_allocations.length, text("Memory leak in TestAllocator. Allocations: ", _allocations));
        }
    }
}

// this shouldn't be needed IMHO and is a bug in std.experimental.allocator that dispose
// doesn't handle 2D arrays correctly
void dispose(A, T)(auto ref A allocator, T[] array) {
    static import std.experimental.allocator;
    import std.traits: isArray, Unqual;

    static if(isArray!T) {
        foreach(ref e; array) {
            dispose(allocator, e);
        }
    }

    alias U = Unqual!T;
    std.experimental.allocator.dispose(allocator, cast(U[])array);
}

XLOPER12 toXlOper(T, A)(T val, ref A allocator) if(is(T == double)) {
    return toXlOper(val);
}

XLOPER12 toXlOper(T)(T val) if(is(T == double)) {
    auto ret = XLOPER12();
    ret.xltype = xltypeNum;
    ret.val.num = val;
    return ret;
}


XLOPER12 toXlOper(T)(T val) if(is(T == string)) {
    return toXlOper(val, xlld.memorymanager.allocator);
}

XLOPER12 toXlOper(T, A)(T val, ref A allocator) if(is(T == string)) {
    import std.utf: byWchar;
    import std.stdio;

    // extra space for the length
    auto wval = cast(wchar*)allocator.allocate((val.length + 1) * wchar.sizeof).ptr;
    wval[0] = cast(wchar)val.length;

    int i = 1;
    foreach(ch; val.byWchar) {
        wval[i++] = ch;
    }

    auto ret = XLOPER12();
    ret.xltype = xltypeStr;
    ret.val.str = cast(XCHAR*)wval;

    return ret;
}


@("toXlOper!string ascii")
@system unittest {
    import std.conv: to;

    const str = "foo";
    auto oper = str.toXlOper;
    scope(exit)FreeXLOper(&oper);

    oper.xltype.shouldEqual(xltypeStr);
    (cast(int)oper.val.str[0]).shouldEqual(str.length);
    (cast(wchar*)oper.val.str)[1 .. str.length + 1].to!string.shouldEqual("foo");
}

@("toXlOper!string allocator")
@system unittest {
    // should throw unless allocations match deallocations
    TestAllocator allocator;
    auto oper = "foo".toXlOper(allocator);
    allocator.numAllocations.shouldEqual(1);
    FreeXLOper(&oper, allocator);
}

XLOPER12 toXlOper(T)(T[][] values) if(is(T == double) || is(T == string))
{
    return toXlOper(values, xlld.memorymanager.allocator);
}

XLOPER12 toXlOper(T, A)(T[][] values, ref A allocator)
    if(is(T == double) || is(T == string))
{
    import std.algorithm: map, all;
    import std.array: array;
    import std.exception: enforce;
    import std.conv: text;

    static const exception = new Exception("# of columns must all be the same and aren't");
    if(!values.all!(a => a.length == values[0].length))
       throw exception;

    auto ret = XLOPER12();
    ret.xltype = xltypeMulti;
    const rows = cast(int)values.length;
    ret.val.array.rows = rows;
    const cols = cast(int)values[0].length;
    ret.val.array.columns = cols;

    ret.val.array.lparray = cast(XLOPER12*)allocator.allocate(rows * cols * ret.sizeof).ptr;
    auto opers = ret.val.array.lparray[0 .. rows*cols];

    int i;
    foreach(ref row; values)
        foreach(ref val; row) {
            opers[i++] = val.toXlOper(allocator);
        }

    return ret;
}


@("toXlOper string[][]")
@system unittest {
    auto oper = [["foo", "bar", "baz"], ["toto", "titi", "quux"]].toXlOper;
    scope(exit) FreeXLOper(&oper);

    oper.xltype.shouldEqual(xltypeMulti);
    oper.val.array.rows.shouldEqual(2);
    oper.val.array.columns.shouldEqual(3);
    auto opers = oper.val.array.lparray[0 .. oper.val.array.rows * oper.val.array.columns];

    opers[0].shouldEqualDlang("foo");
    opers[3].shouldEqualDlang("toto");
    opers[5].shouldEqualDlang("quux");
}

@("toXlOper string[][] allocator")
@system unittest {
    TestAllocator allocator;
    auto oper = [["foo", "bar", "baz"], ["toto", "titi", "quux"]].toXlOper(allocator);
    allocator.numAllocations.shouldEqual(7);
    FreeXLOper(&oper, allocator);
}

@("toXlOper double[][] allocator")
@system unittest {
    TestAllocator allocator;
    auto oper = [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]].toXlOper(allocator);
    allocator.numAllocations.shouldEqual(1);
    FreeXLOper(&oper, allocator);
}


XLOPER12 toXlOper(T)(T values) if(is(T == string[]) || is(T == double[])) {
    return toXlOper(values, xlld.memorymanager.allocator);
}

XLOPER12 toXlOper(T, A)(T values, ref A allocator) if(is(T == string[]) || is(T == double[])) {
    T[1] realValues = [values];
    return realValues.toXlOper(allocator);
}


@("toXlOper string[] allocator")
@system unittest {
    TestAllocator allocator;
    auto oper = ["foo", "bar", "baz", "toto", "titi", "quux"].toXlOper(allocator);
    allocator.numAllocations.shouldEqual(7);
    FreeXLOper(&oper, allocator);
}

auto fromXlOper(T, A)(LPXLOPER12 val, ref A allocator) if(is(T == double)) {
    return fromXlOper!T(val);
}

auto fromXlOper(T)(LPXLOPER12 val) if(is(T == double)) {
    if(val.xltype == xltypeMissing)
        return double.init;

    return val.val.num;
}

@("isNan for fromXlOper!double")
@system unittest {
    import std.math: isNaN;
    XLOPER12 oper;
    oper.xltype = xltypeMissing;
    fromXlOper!double(&oper).isNaN.shouldBeTrue;
}

auto fromXlOper(T)(ref XLOPER12 val) {
    return (&val).fromXlOper!T;
}

auto fromXlOper(T, A)(ref XLOPER12 val, ref A allocator) {
    return (&val).fromXlOper!T(allocator);
}


// 2D slices
auto fromXlOper(T)(LPXLOPER12 val) if(is(T: E[][], E) && (is(E == string) || is(E == double)))
{
    return fromXlOper!T(val, xlld.memorymanager.allocator);
}


auto fromXlOper(T, A)(LPXLOPER12 val, ref A allocator)
    if(is(T: E[][], E) && (is(E == string) || is(E == double)))
{
    return val.fromXlOperMulti!(Dimensions.Two, typeof(T.init[0][0]))(allocator);
}

@("fromXlOper!string[][]")
unittest {
    auto strings = [["foo", "bar", "baz"], ["toto", "titi", "quux"]];
    auto oper = strings.toXlOper;
    scope(exit) FreeXLOper(&oper);
    oper.fromXlOper!(string[][]).shouldEqual(strings);
}

@("fromXlOper!double[][]")
unittest {
    auto doubles = [[1.0, 2.0], [3.0, 4.0]];
    auto oper = doubles.toXlOper;
    scope(exit) FreeXLOper(&oper);
    oper.fromXlOper!(double[][]).shouldEqual(doubles);
}

@("fromXlOper!string[][] allocator")
unittest {
    TestAllocator allocator;
    auto strings = [["foo", "bar", "baz"], ["toto", "titi", "quux"]];
    auto oper = strings.toXlOper(allocator);
    auto backAgain = oper.fromXlOper!(string[][])(allocator);

    allocator.numAllocations.shouldEqual(16);

    FreeXLOper(&oper, allocator);
    backAgain.shouldEqual(strings);
    allocator.dispose(backAgain);
}

@("fromXlOper!double[][] allocator")
unittest {
    TestAllocator allocator;
    auto doubles = [[1.0, 2.0], [3.0, 4.0]];
    auto oper = doubles.toXlOper(allocator);
    auto backAgain = oper.fromXlOper!(double[][])(allocator);

    allocator.numAllocations.shouldEqual(4);

    FreeXLOper(&oper, allocator);
    backAgain.shouldEqual(doubles);
    allocator.dispose(backAgain);
}


private enum Dimensions {
    One,
    Two,
}

// 1D slices
auto fromXlOper(T)(LPXLOPER12 val)
    if(is(T: E[], E) && (is(E == string) || is(E == double)))
{
    return fromXlOper!T(val, xlld.memorymanager.allocator);
}

// 1D slices
auto fromXlOper(T, A)(LPXLOPER12 val, ref A allocator)
    if(is(T: E[], E) && (is(E == string) || is(E == double)))
{
    return val.fromXlOperMulti!(Dimensions.One, typeof(T.init[0]))(allocator);
}


@("fromXlOper!string[]")
unittest {
    auto strings = ["foo", "bar", "baz", "toto", "titi", "quux"];
    auto oper = strings.toXlOper;
    scope(exit) FreeXLOper(&oper);
    oper.fromXlOper!(string[]).shouldEqual(strings);
}

@("fromXlOper!double[]")
unittest {
    auto doubles = [1.0, 2.0, 3.0, 4.0];
    auto oper = doubles.toXlOper;
    scope(exit) FreeXLOper(&oper);
    oper.fromXlOper!(double[]).shouldEqual(doubles);
}

@("fromXlOper!string[] allocator")
unittest {
    TestAllocator allocator;
    auto strings = ["foo", "bar", "baz", "toto", "titi", "quux"];
    auto oper = strings.toXlOper(allocator);
    auto backAgain = oper.fromXlOper!(string[])(allocator);

    allocator.numAllocations.shouldEqual(14);

    backAgain.shouldEqual(strings);
    FreeXLOper(&oper, allocator);
    allocator.dispose(backAgain);
}

@("fromXlOper!double[] allocator")
unittest {
    TestAllocator allocator;
    auto doubles = [1.0, 2.0, 3.0, 4.0];
    auto oper = doubles.toXlOper(allocator);
    auto backAgain = oper.fromXlOper!(double[])(allocator);

    allocator.numAllocations.shouldEqual(2);

    backAgain.shouldEqual(doubles);
    FreeXLOper(&oper, allocator);
    allocator.dispose(backAgain);
}


private auto fromXlOperMulti(Dimensions dim, T, A)(LPXLOPER12 val, ref A allocator) {
    import xlld.xl: coerce, free;
    import std.exception: enforce;
    import std.experimental.allocator: makeArray;

    static const exception = new Exception("XL oper not of multi type");

    const realType = val.xltype & ~xlbitDLLFree;
    if(realType != xltypeMulti)
        throw exception;

    const rows = val.val.array.rows;
    const cols = val.val.array.columns;

    static if(dim == Dimensions.Two) {
        auto ret = allocator.makeArray!(T[])(rows);
        foreach(ref row; ret)
            row = allocator.makeArray!T(cols);
    } else static if(dim == Dimensions.One) {

        auto ret = allocator.makeArray!T(rows * cols);
    } else
        static assert(0);

    auto values = val.val.array.lparray[0 .. (rows * cols)];

    foreach(const row; 0 .. rows) {
        foreach(const col; 0 .. cols) {
            auto cellVal = coerce(&values[row * cols + col]);
            scope(exit) free(&cellVal);

            auto value = cellVal.xltype == dlangToXlOperType!T.Type ? cellVal.fromXlOper!T(allocator) : T.init;
            static if(dim == Dimensions.Two)
                ret[row][col] = value;
            else
                ret[row * cols + col] = value;
        }
    }

    return ret;
}


auto fromXlOper(T)(LPXLOPER12 val) if(is(T == string)) {
    return fromXlOper!T(val, xlld.memorymanager.allocator);
}

auto fromXlOper(T, A)(LPXLOPER12 val, ref A allocator) if(is(T == string)) {

    import std.experimental.allocator: makeArray;
    import std.utf;

    if(val.xltype == xltypeMissing)
        return null;

    auto ret = allocator.makeArray!char(val.val.str[0]);
    int i;
    foreach(ch; val.val.str[1 .. ret.length + 1].byChar)
        ret[i++] = ch;

    return cast(string)ret;
}

@("fromXlOper missing")
@system unittest {
    XLOPER12 oper;
    oper.xltype = xltypeMissing;
    fromXlOper!string(&oper).shouldBeNull;
}

@("fromXlOper string allocator")
@system unittest {
    TestAllocator allocator;
    auto oper = "foo".toXlOper(allocator);
    auto str = fromXlOper!string(&oper, allocator);
    allocator.numAllocations.shouldEqual(2);

    FreeXLOper(&oper, allocator);
    str.shouldEqual("foo");
    allocator.dispose(cast(void[])str);
}

private enum isWorksheetFunction(alias F) =
    isSupportedFunction!(F, double, double[][], string[][], string[], double[], string);

@safe pure unittest {
    import xlld.test_d_funcs;
    // the line below checks that the code still compiles even with a private function
    // it might stop compiling in a future version when the deprecation rules for
    // visibility kick in
    static assert(!isWorksheetFunction!shouldNotBeAProblem);
    static assert(!isWorksheetFunction!FuncThrows);
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

@("Only look at nothrow functions")
@system unittest {
    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");
    auto arg = toXlOper(2.0);
    static assert(!__traits(compiles, FuncThrows(&arg)));
}

@("FuncAddEverything wrapper is @nogc")
@system @nogc unittest {
    import std.experimental.allocator.mallocator: Mallocator;

    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");
    auto arg = toXlOper(2.0, Mallocator.instance);
    scope(exit) FreeXLOper(&arg, Mallocator.instance);
    FuncAddEverything(&arg);
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

/**
 A string to use with `mixin` that wraps a D function
 */
string wrapModuleFunctionStr(string moduleName, string funcName)() {
    if(!__ctfe) {
        return "";
    }

    import std.array: join;
    import std.traits: Parameters, functionAttributes, FunctionAttribute;
    import std.conv: to;
    import std.algorithm: map;
    import std.range: iota;
    mixin("import " ~ moduleName ~ ": " ~ funcName ~ ";");

    const argsLength = Parameters!(mixin(funcName)).length;
    // e.g. LPXLOPER12 arg0, LPXLOPER12 arg1, ...
    const argsDecl = argsLength.iota.map!(a => `LPXLOPER12 arg` ~ a.to!string).join(", ");
    // e.g. arg0, arg1, ...
    const argsCall = argsLength.iota.map!(a => `arg` ~ a.to!string).join(", ");
    const nogc = functionAttributes!(mixin(funcName)) & FunctionAttribute.nogc
        ? "@nogc"
        : "";
    return [
        `extern(Windows) LPXLOPER12 ` ~ funcName ~ `(` ~ argsDecl ~ `) nothrow ` ~ nogc ~ `{`,
        `    static import ` ~ moduleName ~ `;`,
        `    alias wrappedFunc = ` ~ moduleName ~ `.` ~ funcName ~ `;`,
        `    return wrapModuleFunctionImpl!wrappedFunc(` ~ argsCall ~  `);`,
        `}`,
    ].join("\n");
}

/**
 Implemented a wrapper for a regular D function
 */
LPXLOPER12 wrapModuleFunctionImpl(alias wrappedFunc, T...)(T args) {
    return wrapModuleFunctionImplAllocator!wrappedFunc(xlld.memorymanager.allocator, args);
}

LPXLOPER12 wrapModuleFunctionImplAllocator(alias wrappedFunc, A, T...)(ref A allocator, T args) {
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
        try {
            dArgs[i] = fromXlOper!InputType(&realArgs[i], allocator);
        } catch(Exception ex) {
            ret.xltype = xltypeErr;
            ret.val.err = -1;
            return &ret;
        }
    }

    try
        ret = toXlOper(wrappedFunc(dArgs.expand), allocator);
    catch(Exception ex)
        return null;

    foreach(ref dArg; dArgs) {
        allocator.dispose(dArg);
    }

    ret.xltype |= xlbitDLLFree;

    return &ret;
}

@("No memory allocation bugs in wrapModuleFunctionImplAllocator for double return")
@system unittest {
    import xlld.test_d_funcs: FuncAddEverything;

    TestAllocator allocator;
    auto arg = toSRef([1.0, 2.0]);
    auto oper = wrapModuleFunctionImplAllocator!FuncAddEverything(allocator, &arg);
    (oper.xltype & xlbitDLLFree).shouldBeTrue;
    allocator.numAllocations.shouldEqual(2);
    oper.shouldEqualDlang(3.0);
    FreeXLOper(oper, allocator); // normally this is done by Excel
}

@("No memory allocation bugs in wrapModuleFunctionImplAllocator for double[][] return")
@system unittest {
    import xlld.test_d_funcs: FuncTripleEverything;

    TestAllocator allocator;
    auto arg = toSRef([1.0, 2.0, 3.0]);
    auto oper = wrapModuleFunctionImplAllocator!FuncTripleEverything(allocator, &arg);
    (oper.xltype & xlbitDLLFree).shouldBeTrue;
    (oper.xltype & ~xlbitDLLFree).shouldEqual(xltypeMulti);
    allocator.numAllocations.shouldEqual(3);
    oper.shouldEqualDlang([[3.0, 6.0, 9.0]]);
    FreeXLOper(oper, allocator); // normally this is done by Excel
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

    static exception = new const Exception("Error converting input");

    if(arg.xltype != dlangToXlOperType!T.InputType)
        throw exception;

    auto realArg = coerce(arg);

    if(realArg.xltype != dlangToXlOperType!T.Type) {
        free(&realArg);
        throw exception;
    }

    return realArg;
}
