/**
 Enables linking on  Windows without having to link to the real implementations
 Only for unit testing.
 */
module xlld.dummy;

version(unittest):
version(Windows):

import xlld.xlcall;

extern(System) int Excel4v(int xlfn, LPXLOPER operRes, int count, LPXLOPER* opers) { //pascal
    return 0;
}

extern(C) int Excel4(int xlfn, LPXLOPER operRes, int count,... ) { //_cdecl
    return 0;
}
