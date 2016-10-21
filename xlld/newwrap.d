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

string wrapWorksheetFunctionsString(string module_)() {
    import std.array: join;
    enum funcName = `FuncAddEverything`;
    return
        [
            `static import ` ~ module_ ~ `;`,
            `extern(Windows) double ` ~ funcName ~ `(LPXLOPER12 arg) {`,
            `    return ` ~ module_ ~ `.` ~ funcName ~ `(arg.fromXlOper!(double[][]));`
            `}`,
        ].join("\n");
}

@("Wrap test_d_funcs.FuncAddEverything")
@system unittest {
    import std.algorithm;
    import std.array;

    mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");

    auto arg = XLOPER12();
    arg.xltype = xltypeMulti;
    arg.val.array.rows = 2;
    arg.val.array.columns = 4;

    double[] values = [1, 2, 3, 4, 11, 12, 13, 14];
    auto operValues = values.map!(a => a.toXlOper).array;
    arg.val.array.lparray = operValues.ptr;

    //auto expected = values.fold!((a, b) => a + b)(0.0).toXlOper;
    auto expected = values.fold!((a, b) => a + b)(0.0);
    FuncAddEverything(&arg).shouldEqual(expected);

    values = [0, 1, 2, 3, 4, 5, 6, 7];
    operValues = values.map!(a => a.toXlOper).array;
    arg.val.array.lparray = operValues.ptr;

    expected = values.fold!((a, b) => a + b)(0.0);
    FuncAddEverything(&arg).shouldEqual(expected);
}
