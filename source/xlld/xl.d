/**
 Wraps calls to xlXXX "functions" via the Excel4/Excel12 functions
 */
module xlld.xl;

import xlld;

version(unittest) {
    XLOPER12 coerce(LPXLOPER12 oper) {

    }

    void free(LPXOPER12 oper) {

    }
} else {
    XLOPER12 coerce(LPXLOPER12 oper) {
        XLOPER12 coerced;
        Excel12f(xlCoerce, &coerced, [oper]);
        return coerced;
    }

    void free(LPXLOPER12 oper) {
        Excel12f(xlFree, null, [oper]);
    }
}
