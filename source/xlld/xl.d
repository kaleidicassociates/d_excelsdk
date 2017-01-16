/**
 Wraps calls to xlXXX "functions" via the Excel4/Excel12 functions
 */
module xlld.xl;

import xlld.xlcall;
import xlld.wrap;

version(unittest) {

    import xlld.xlcallcpp: EXCEL12PROC, SetExcel12EntryPt;

    static this() {
        SetExcel12EntryPt(&excel12UnitTest);
    }

    extern(Windows) int excel12UnitTest (int xlfn, int numOpers, LPXLOPER12 *opers, LPXLOPER12 result) nothrow @nogc {
        switch(xlfn) {

        default:
            return xlretFailed;

        case xlFree:
            return xlretSuccess;

        case xlCoerce:
            assert(numOpers == 1);
            auto oper = opers[0];
            *result = *oper;

            switch(oper.xltype) {

            case xltypeSRef:
                result.xltype = gReferencedType;
                break;

            case xltypeNum:
            case xltypeStr:
                result.xltype = oper.xltype;
                break;

            default:
            }

            return xlretSuccess;
        }
    }
}

XLOPER12 coerce(LPXLOPER12 oper) nothrow @nogc {
    import xlld.framework: Excel12f;

    XLOPER12 coerced;
    LPXLOPER12[1] arg = [oper];
    Excel12f(xlCoerce, &coerced, arg);
    return coerced;
}

void free(LPXLOPER12 oper) nothrow @nogc {
    import xlld.framework: Excel12f;

    LPXLOPER12[1] arg = [oper];
    Excel12f(xlFree, null, arg);
}
