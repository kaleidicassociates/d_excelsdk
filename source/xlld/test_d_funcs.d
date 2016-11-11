/**
 Only exists to test the wrapping functionality Contains functions
 with regular D types that will get wrapped so they can be called by
 the spreadsheet.
 */
module xlld.test_d_funcs;

double FuncAddEverything(double[][] args) {
    import std.algorithm: fold;
    import std.math: isNaN;

    double ret = 0;
    foreach(row; args)
        ret += row.fold!((a, b) => b.isNaN ? 0.0 : a + b)(0.0);
    return ret;
}

double[][] FuncTripleEverything(double[][] args) {
    double[][] ret;
    ret.length = args.length;
    foreach(i; 0 .. args.length) {
        ret[i].length = args[i].length;
        foreach(j; 0 .. args[i].length)
            ret[i][j] = args[i][j] * 3;
    }

    return ret;
}

double FuncAllLengths(string[][] args) {
    import std.algorithm: fold;

    double ret = 0;
    foreach(row; args)
        ret += row.fold!((a, b) => a + b.length)(0.0);
    return ret;
}

double[][] FuncLengths(string[][] args) {
    double[][] ret;

    ret.length = args.length;
    foreach(i; 0 .. args.length) {
        ret[i].length = args[i].length;
        foreach(j; 0 .. args[i].length)
            ret[i][j] = args[i][j].length;
    }

    return ret;
}


string[][] FuncBob(string[][] args) {
    string[][] ret;

    ret.length = args.length;
    foreach(i; 0 .. args.length) {
        ret[i].length = args[i].length;
        foreach(j; 0 .. args[i].length)
            ret[i][j] = args[i][j] ~ "bob";
    }

    return ret;
}


double FuncDoubleSlice(double[] arg) {
    return arg.length;
}

double FuncStringSlice(string[] arg) {
    return arg.length;
}

double[] FuncSliceTimes3(double[] arg) {
    import std.algorithm;
    import std.array;
    return arg.map!(a => a * 3).array;
}

string[] StringsToStrings(string[] args) {
    import std.algorithm;
    import std.array;
    return args.map!(a => a ~ "foo").array;
}

string StringsToString(string[] args) {
    import std.string;
    return args.join(", ");
}