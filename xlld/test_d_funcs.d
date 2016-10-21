/**
 Only exists to test the wrapping functionality Contains functions
 with regular D types that will get wrapped so they can be called by
 the spreadsheet.
 */
module xlld.test_d_funcs;

version(unittest):

import unit_threaded;

double FuncAddEverything(double[][] args) {
    import std.algorithm: fold;

    double ret = 0;
    foreach(row; args)
        ret += row.fold!((a, b) => a + b)(0.0);
    return ret;
}

unittest {
    FuncAddEverything([[1, 2], [3, 4]]).shouldEqual(10);
    FuncAddEverything([]).shouldEqual(0);
}
