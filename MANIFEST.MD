Framework.d
===========
Translated from SDK framework.c
                debugPrintf
                GetTempMemory
                FreeAllTempMemory
                Excel
                Excel12f
                TempNum
                TempNum12
                TempStr
                TempStrConst
                TempStr12
                TempBool
                TempBool12
                TempInt
                TempInt12
                TempErr
                TempErr12
                TempActiveRef
                TempActiveRef12
                TempActiveCell
                TempActiveCell12
                TempActiveRow
                TempActiveRow12
                TempActiveColumn
                TempActiveColumn12
                TempMissing
                TempMissing12
                InitFramework
		QuitFramework

Generic.d
=========
Translated from generic.c and generic.h
Template for creating XLLs for Microsoft Excel.
  
                This file contains sample code you can use as 
                a template for writing your own Microsoft Excel XLLs. 
                An XLL is a DLL that stands alone, that is, you
                can open it by choosing the Open command from the
                File menu. This code demonstrates many of the features 
                of the Microsoft Excel C API.
                
                DllMain
                xlAutoOpen
                xlAutoClose
                lpstricmp
                xlAutoRegister12
                xlAutoAdd
                xlAutoRemove
                xlAddInManagerInfo12
                DIALOGMsgProc
                ExcelCursorProc
                HookExcelWindow
                UnhookExcelWindow
                fShowDialog
                GetHwnd
                Func1
                FuncSum
                fDance
                fDialog
                fExit
                FuncFib

Generic32.def
=============
DLL export definition file for 32 bit
                xlAutoOpen                      @2
                xlAutoClose                     @3
                xlAutoRegister12        @4
                xlAutoAdd                       @5
                xlAutoRemove            @6
                xlAddInManagerInfo12 @7
                DIALOGMsgProc           @8
                ExcelCursorProc @9
                HookExcelWindow @10
                UnhookExcelWindow       @11
                fShowDialog     @12
                GetHwnd @13
                Func1   @14
                FuncSum @15
                fDance  @16
                fDialog @17
                fExit   @18
                FuncFib @19
                WrapSquare3 @20
                lastErrorMessage        @21

Generic64.def
=============
DLL export definitions for 64 bit

     xlAutoOpen			@2
                xlAutoClose			@3
                xlAutoRegister12	@4
                xlAutoAdd			@5
                xlAutoRemove		@6
                xlAddInManagerInfo12 @7
                DIALOGMsgProc		@8
                ExcelCursorProc	@9
                HookExcelWindow	@10
                UnhookExcelWindow	@11
                fShowDialog	@12
                GetHwnd	@13
                Func1	@14
                FuncSum	@15
                fDance	@16
				fDialog @17
                fExit	@18
                FuncFib	@19
                WrapSquare3 @20
                lastErrorMessage	@21

Make32.bat
==========
Sample make file to create a 32 bit DLL

Make64.bat
==========
Sample make file to create a 64 bit DLL

MemoryManager.D
===============
Ported from MemoryManager.cpp.  This struct provides
//              each thread with an array of bytes to use as temporary memory.
//              The size of the array, and the methods for dealing with the
//              memory explicitly, is in the class MemoryPool


MemoryPool.d
=============
Ported from MemoryPool.cpp and MemoryPool.h
A memory pool is an array of characters that is pre-allocated,
//              and used as temporary memory by the caller. The allocation
//              algorithm is very simple. When a thread asks for some memory,
//              the index into the array moves forward by that many bytes, and
//              a pointer is returned to the previous index before the pointer
//              was advanced. When a call comes to free all of the memory, the
//              pointer is set back to the beginning of the array.

Ole2.d
======
Standard OLE header

Wrap.d
======
Idiomatic D wrapper for Excel (first version) - mostly wraps Excel types and converts to and from D types.  No UDA magic
        struct ExcelResult(T)
        enum ExcelReturnStatus

        auto fromXLOPER12(T=Variant[][])(LPXLOPER12 pxArg)
        if (is(T==Variant[][]) || (is(T==double[][])) || is(T==double[]))

        struct excelCallPool
        
        extern(Windows) void xlAutoFree12(LPXLOPER12 p)

        LPXLOPER12 makeXLOPER12(Variant[] arg)

        LPXLOPER12 makeXLOPER12(Variant[][] arg)

        LPXLOPER12 makeXLOPER12(T)(T[][] arg)
        if (is(T==double))
        
        LPXLOPER12 makeXLOPER12(T)(T[] arg)
        if (isSomeString!(T))
        
        LPXLOPER12 makeXLOPER12(T)(T[][] arg)
        if (isSomeString!(T))
        
        LPXLOPER12 makeXLOPER12Error(int error=-1)

string makeMultiArgWrap(string wrapperName, int numArgs)

Xlcall.d
========
Port of xlcall.h
defines constants and data types used in Excel C API

xlcallcpp.d
============
Port of xlcall.cpp
defines entry points used in Excel C API

xllapi.d
========
not sure - possibly Stefan's reworking of existing C API from other files

xlltypes.d
==========
probably Stefan's reworking of types for existing C API

wrap/boilerplate.d
==================

        DllMain
        xlAutoOpen() - automatic registration of functions from static array
        xlAutoClose() - automatic deregistation of functions
        exitDLL() - DLL triggered exit - unregisters each function
        xlAutoRegister12
        xlAutoRemove()
        xlAddInManagerInfo12()
        GetHwnd

wrap/dispatch.d
===============
COM wrapper?

wrap/frameworkallocator.d
==========================
replacing standard MS version with Andrei's allocator

wrap/helper.d
=============
lpwstricmp - compare pascal string and C string insensitive
        DIALOGMsgProc - process dialogue message
        ExcelCursorProc - change cursor type from busy to normal
        HookExcelWindow
        UnhookExcelWindow
        GetHwnd

wrap/lazy.d
===========
Stefan's work on turning UDAs into registration

wrap/memorymanager.d
====================
memory manager - how does this fit with wrap/frameworkallocator.d ?

wrap/traits.d
=============
Stefan's work on turning UDAs into registration

wrap/traitshelper.d
===================
Stefan's work on turning UDAs into registration

wrap/traitstest.d
=================
Sample use of UDAs

wrap/vbahelpers.d
=================
ignore this for now.

wrap/xloper.d
=============
same as wrap.d in base directory

        

        






