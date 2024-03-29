/**
	Idiomatic D wrapper for Excel SDK inspired by PyD and PyXLL:
		- allows writing excel user defined functions in D without extra glue code
		- simply annotate the function with xlld attributes and it will be wrapped and automatically registered
		  on workbook open and deregistered on close

	That's the ultimate goal.  For the time being we start by using annotations to automatically register and
	deregister functions but still require user to accept LPXLOPER as arguments and return type.  Function should be
	extern(Windows) too.  Later we can transparently wrap these as D variants/native types.

	To do:
		- wrap arguments and return types
		- custom menu helper
		- custom toolbar helper
		- dialogue box helper
		- handle more types: int, long, float, double, string, bool, DateTime/SysTime, D matrices
		- excel errors: null, div/0, value, ref, name, num,na (map to exceptions)
			Lookup, ZeroDivision, Value, Reference, Name, Arithmetic, Runtime
		- cell metadata: value, address, formula, note, sheetName, sheetId, rect (xlRect - coordinates)
		- async returns
		- calling excel from D
		- COM
		- callbacks: onOpen, onReload, onClose, reloadXLL,
		- object cache
*/


/*
	Don't forget to figure out when to set the xlAutoFree bits!
*/

module xlld.wrap;
import std.variant;
import win32.winuser:PostMessage,CallWindowProc,GetWindowLongPtr,SetWindowLongPtr,DialogBox;
import std.c.windows.windows;
import core.sys.windows.windows;
import xlcall;
import framework;
import core.stdc.wchar_ : wcslen;
import core.stdc.wctype:towlower;
import std.format;
  
/+
  
/**
   Syntax of the Register Command:
        REGISTER(module_text, procedure, type_text, function_text, 
                 argument_text, macro_type, category, shortcut_text,
                 help_topic, function_help, argument_help1, argument_help2,...)
  
  
   g_rgWorksheetFuncs will use only the first 11 arguments of 
   the Register function.
  
   This is a table of all the worksheet functions exported by this module.
   These functions are all registered (in xlAutoOpen) when you
   open the XLL. Before every string, leave a space for the
   byte count. The format of this table is the same as 
   arguments two through eleven of the REGISTER function.
   g_rgWorksheetFuncsRows define the number of rows in the table. The
   g_rgWorksheetFuncsCols represents the number of columns in the table.
*/
enum g_rgWorksheetFuncsRows =3;
enum g_rgWorksheetFuncsCols =10;

__gshared wstring[g_rgWorksheetFuncsCols][g_rgWorksheetFuncsRows] exportedFunctionTable =
[
	[ "Func1"w,                                     // Procedure
		"UU"w,                                  // type_text
		"Func1"w,                               // function_text
		"Arg"w,                                 // argument_text
		"1"w,                                   // macro_type
		"Generic Add-In"w,                      // category
		""w,                                    // shortcut_text
		""w,                                    // help_topic
		"Always returns the string 'Func1'"w,   // function_help
		"Argument ignored"w                     // argument_help1
	],
	[ "FuncSum"w,
		"UUUUUUUUUUUUUUUUUUUUUUUUUUUUUU"w, // up to 255 args in Excel 2007 and later,
										   // upto 29 args in Excel 2003 and earlier versions
		"FuncSum"w,
		"number1,number2,..."w,
		"1"w,
		"Generic Add-In"w,
		""w,                                    
		""w,                                  
		"Adds the arguments"w,   
		"Number1,number2,... are 1 to 29 arguments for which you want to sum."w                   
	],
	[ "FuncFib"w,
		"UU"w,	
		"FuncFib"w,
		"Compute to..."w,
		"1"w,
		"Generic Add-In"w,
		""w,
		""w,
		"Number to compute to"w
		"Computes the nth fibonacci number"w,
	],
];

/*

	0	function name
	1	type text (up to 29 args excel 2003 and prev; up to 255 args excel 2007 and later)
	2	function text
	3	argument text eg "number1,number2,..."
	4	macro type = 1
	5	category for add-in
	6	shortcut text - can be blank
	7	help topic - can be blank
	8	function help eg "returns the hash of arguments"
	9	argument help1 eg "number1,2,... are 1 to 29 arguments to process"
*/

struct xllCategory(string _name)
{
	enum bool IsPyName = T[0].stringof.startsWith("xllCategory!");
}

@xlld!(	xllRename!"newname",xllAtLeastArgs!(2),xllCategory!"dlangsci",xllShortcut!"shortcutName",xllHelpTopic!"general",
	xllFunctionHelp!"returns the hash of arguments",xllArgumentHelp!"number 1,2..29 are 1 to 29 arguments to process",
	xllThreadSafe!false,xllMacro!false,	xllAllowAbort!false,xllVolatile!false,xllDisableFunctionWizard!false,
	xlldisableReplaceCalc!false)
+/

/**
	To do:
		- toolbar construction helper
		- dialogue box construction helper
*/



struct ExcelResult(T)
{
	bool success;
	ExcelReturnStatus status;
	T[][] data;
	alias data this;
}

enum ExcelReturnStatus
{
	success,
	missing,
	wrongType,
	wrongShape,
	excelError,
	unhandledType,
	uncalced,
}


auto fromXLOPER12(T=Variant[][])(LPXLOPER12 pxArg)
if (is(T==Variant[][]) || (is(T==double[][])) || is(T==double[]))
{
	ExcelResult!T ret;
	long i;					// Row and column counters for arrays 
	LPXLOPER12 px;			// Pointer into array 
	ret.success=true;
	ret.status=ExcelReturnStatus.success;

	switch (pxArg.xltype)
	{	
		case xltypeMissing:
			return ExcelResult(false,ExcelReturnStatus.missing);

		case xltypeNum:
			static if (is(T==double[]))
				return ExcelResult(true,ExcelReturnStatus.success,[pxArg.val.num]);
			else
				return ExcelResult(true,ExcelReturnStatus.success,[[pxArg.val.num]]);
			break;

		case xltypeStr:
			static if(is(T==(double[])))
				return ExcelDoubleResultVector(false,ExcelReturnStatus.wrongType,[double.nan]);
			else static if(is(T==double[][]))
				return ExcelDoubleResultMatrix(false,ExcelReturnStatus.wrongType,[[double.nan]]);
			else
				return ExcelVariantResultMatrix(true,ExcelReturnStatus.success,[[pxArg.val.str.to!string()]]);

		case xltypeRef:
		case xltypeSRef:
		case xltypeMulti:
			XLOPER12 xMulti;		// Argument coerced to xltypeMulti 
			scope(exit)
			{
				Excel12f(xlFree, cast(XLOPER12*)0, [cast(LPXLOPER12) &xMulti]);
				Excel12f(xlFree, cast(XLOPER12*)0, [cast(LPXLOPER12) pxArg]);
			}
			//	coerce might have failed due to an uncalced cell
			// Microsoft Excel will call us again in a moment after that cell has been calced.
			if (xlretUncalced == Excel12f(xlCoerce, &xMulti, [cast(LPXLOPER12) pxArg, TempInt12(xltypeMulti)]))
			{
				static if(is(T==Variant[][]))
					return ExcelVariantResultMatrix(false,ExcelReturnStatus.uncalced,[[double.nan]]);
				else static if(is(T==double[][]))
					return ExcelDoubleResultMatrix(false,ExcelReturnStatus.uncalced,[[double.nan]]);
				else static if(is(T==double[]))
					return ExcelDoubleResultVector(false,ExcelReturnStatus.uncalced,[[double.nan]]);
			}
			auto numRows=xMulti.val.array.rows;
			auto numCols=xMulti.val.array.columns;
			static if(is(T==double[]))
			{
				if (!(numRows==1 || numCols==1))
					return ExcelDoubleResultVector(false,ExcelReturnStatus.wrongShape,[double.nan]);
				bool flipped=(numCols!=1);
				if (!flipped)
					ret.length=numRows.length;
				else
					ret.length=numCols.length;
			}
			else static if(is(T==Variant[][]) || is(T==double[][]))
			{
				ret.length=numRows;
				foreach(row;ret)
					row.length=numCols;
			}
			foreach(i;0..numRows)
			{
				foreach(j;0..numCols)
				{
					// obtain a pointer to the current item //
					px = xMulti.val.array.lparray(i*numCols+j);
					// switch on XLOPER12 type //
					switch (px.xltype)
					{
						case xltypeNum:
							static if ((is(T==Variant[][])) || (is(T==double[][])))
								ret[i][j]= px.val.num;
							else // must be double[]
							{
								if(!flipped)
									ret[i]=px.val.num;
								else
									ret[j]=px.val.num;
							}
							break;

						// if an error store in error //
						case xltypeErr::
							status=ExcelReturnStatus.excelError;
							static if ((is(T==Variant[][]))||(is(T==double[][])))
							{
								ret[i][j]=px.val.err;
							}
							else // double[]
							{
								if (!flipped)
									ret[i]=px.val.err;
								else
									ret[j]=px.val.err;
							}
							break;

						case xltypeNil:
							static if ((is(T==Variant[][])))
								ret[i][j]=null;
							else static if((is(T==double[][])))
								ret[i][j]=double.nan;
							else // double[]
								ret[i]=double.nan;
							else
								ret[j]=double.nan;
							break;

						// if anything else set error //
						default:
							static if(is(T==Variant[][]))
								return ExcelVariantResultMatrix(false,ExcelReturnStatus.unhandledType,[[double.nan]]);
							else static if(is(T==double[][]))
								return ExcelVariantResultMatrix(false,ExcelReturnStatus.unhandledType,[[double.nan]]);
							else
								return ExcelVariantResultVector(false,ExcelReturnStatus.unhandledType,[double.nan]);
					}
				}
			}
			ret.success=(status==ExcelReturnStatus.success);
			ret.status=status;
			return ret;
			static if(is(T))
			// free the returned array //
			break;

		case xltypeErr:
			success=false;
			break;

		default:
			success=false;
			break;
	}
	ret.success=(ret.status==ExcelReturnStatus.success);
	return ret;
}


LPXLOPER12 makeXLOPER12(double arg)
{
	LPXLOPER12 lpx;
	lpx=cast(LPXLOPER12) excelCallPool.allocate(XLOPER12.sizeof);
	lpx.xltype=xltypeNum;
	lpx.val.num=arg;
	return lpx;
}
// need to have allocator!
LPXLOPER12 makeXLOPER12(Flag!"autoFree" autoFree=Flag.autoFree.No)(string arg)
{
	LPXLOPER12 lpx;
	static if(Flag.autoFree.yes)
		lpx = cast(LPXLOPER12) excelCallPool.allocate(XLOPER12.sizeof);
	else
		lpx=theAllocator.allocate.make!XLOPER12(1);
	lpx.xltype = xltypeStr;
	lpx.val.str = arg.makePascalString;
	return lpx;
}
 
LPXLOPER12 makeXLOPER12(Flag!"autoFree" autoFree=Flag.autoFree.No)(Variant[] arg)
{
	return makeXLOPER12!autoFree([arg]);
}

LPXLOPER12 makeXLOPER12(Flag!"autoFree" autoFree=Flag.autoFree.No)(Variant[][] arg)
{
	auto numRows=arg.length;
	if(numRows==0)
	{

	}
	auto numCols=arg[0].length;
	foreach(row;arg[1..$])
		enforce(row.length==numCols, new Exception("makeXLOPER12: arg must be rectangular"));
	
	auto xlValuesLength=numRows*numCols;
	XLOPER12* ret=allocXLOPER12!autoFree(1);
	XLOPER12[] xlValues;
	{
		static if(autoFree.No)
			scope(fail)
				theAllocator.dispose(ret);
 		xlValues= allocXLOPER12Slice!autoFree(xlValuesLength);
 	}
	ret.xltype=xltypeMulti;

	foreach(i;0..numRows)		
	{
		foreach(j;0..numCols)
		{
			if(arg[i][j].convertsTo!(double))
			{
				xlValues[i*numCols+j].val.num=arg[i][j].get!(double);
				xlValues[i*numCols+j].xltype=xltypeNum;
			}
			else if(arg[i][j].convertsTo!(string))
			{
				xlValues[i*numCols+j].val.str=arg[i][j].get!(string).makePascalString;
				xlValues[i*numCols+j].xltype=xltypeStr;
			}
			else if(arg[i][j]==null)
			{
				//xlValues[i*numCols+j].val.
				xlValues[i*numCols+j].xltype=xltypeNil;
			}
			else
			{
				throw new Exception("makeXLOPER12: unknown type row: "~i.to!string~"; col "~j.to!string~" type: "~arg[i][j].typeid.to!string);
			}
		}
	}

	ret.xltype = xltypeMulti;
	ret.val.array.lparray = xlValues.ptr;
	ret.val.array.rows = numRows;
	ret.val.array.columns = numCols;
	return ret;
}

LPXLOPER12 makeXLOPER12(Flag!"autoFree" autoFree=Flag.autoFree.No,T)(T[] arg)
if (isNumeric!(T))
{
	return makeXLOPER12([arg]);
}

private XLOPER12* allocXLOPER12(Flag!"autoFree" autoFree=Flag.autoFree.No)(size_t n=1)
{
	XLOPER12* ret;
	static if(Flag.autoFree.yes)
	{
		ret = cast(XLOPER12*) excelCallPool.allocate(XLOPER12.sizeof*n);
		if (ret is null)
			throw new Exception("allocXLOPER12 memory allocation error for "~n.to!string~" XLOPER12s");
	}
	else
	{
		ret = cast(XLOPER12*) theAllocator.allocate.make!XLOPER12(n);
		if (ret is null)
			throw new Exception("allocXLOPER12 memory allocation error for "~n.to!string~" XLOPER12s");
	}
	return ret;
}

private XLOPER12[] allocXLOPER12Slice(Flag!"autoFree" autoFree=Flag.autoFree.No)(size_t n=1)
{
	XLOPER12[] ret;
	static if(Flag.autoFree.yes)
	{
		ret= cast(XLOPER12[0..xlValuesLength]) excelCallPool.allocate(XLOPER12.sizeof*n);
		if ((ret is null) || (ret.length!=n))
			throw new Exception("allocXLOPER12Slice memory allocation error for "~n.to!string~" XLOPER12s");
	}
	else
	{
		ret=theAllocator.makeArray!XLOPER12(xlValuesLength);
		if (ret.length!=n)
		{
			throw new Exception("allocXLOPER12Slice memory allocation error for "~n.to!string~" XLOPER12s");
		}
	}
}


LPXLOPER12 makeXLOPER12(Flag!"autoFree" autoFree=Flag.autoFree.No,T)(T[][] arg)
if (isNumeric!(T))
{
	auto numCols=arg[0].length;
	foreach(row;arg[1..$])
		enforce(row.length==numCols, new Exception("makeXLOPER12: arg must be rectangular");

	auto xlValuesLength=numRows*numCols;
	XLOPER12* ret=allocXLOPER12!autoFree(1);
	XLOPER12[] xlValues;
	{
		static if(autoFree.No)
			scope(fail)
				theAllocator.dispose(ret);
 		xlValues= allocXLOPER12Slice!autoFree(xlValuesLength);
 	}
	ret.xltype=xltypeMulti;
	auto numRows=arg.length;
	if(numRows==0)
	{

	}
	foreach(i;0..numRows)		
	{
		foreach(j;0..numCols)
		{
			xlValues[i*numCols+j].val.num=arg[i][j];
			xlValues[i*numCols+j].xltype=xltypeNum;
		}
	}

	ret.xltype = xltypeMulti;
	ret.val.array.lparray = xlValues.ptr;
	ret.val.array.rows = numRows;
	ret.val.array.columns = numCols;
	return ret;
}

LPXLOPER12 makeXLOPER12(Flag!"autoFree" autoFree=Flag.autoFree.No,T)(T[] arg)
if (isSomeString!(T))
{
	return makeXLOPER12([arg]);
}

LPXLOPER12 makeXLOPER12(Flag!"autoFree" autoFree=Flag.autoFree.No,T)(T[][] arg)
if (isSomeString!(T))
{
	ret.xltype=xltypeMulti;
	auto numRows=arg.length;
	if(numRows==0)
	{

	}
	auto numCols=arg[0].length;
	foreach(row;arg[1..$])
		enforce(row.length==numCols, new Exception("makeXLOPER12: arg must be rectangular");

	auto xlValuesLength=numRows*numCols;
	XLOPER12* ret=allocXLOPER12!autoFree(1);
	XLOPER12[] xlValues;
	{
		static if(autoFree.No)
			scope(fail)
				theAllocator.dispose(ret);
 		xlValues= allocXLOPER12Slice!autoFree(xlValuesLength);
 	}

	foreach(i;0..numRows)		
	{
		foreach(j;0..numCols)
		{
			xlValues[i*numCols+j].val.str=arg[i][j].to!(wstring).makePascalString!autoFree;
			xlValues[i*numCols+j].xltype=xltypeStr;
		}
	}

	ret.xltype = xltypeMulti;
	ret.val.array.lparray = xlValues;//&xlValues[0];
	ret.val.array.rows = numRows;
	ret.val.array.columns = numCols;
	return cast(LPXLOPER12) &ret;
}

LPXLOPER12 makeXLOPER12Error(Flag!"autoFree" autoFree=Flag.autoFree.No)()
{
	XLOPER12* ret=allocXLOPER12!autoFree(1);
	ret.xltype = xltypeErr;
	ret.val.err = error;
	return cast(LPXLOPER12)ret;
}


string makeMultiArgWrap(string wrapperName, int numArgs)
{
	string ret;
	ret="extern(Windows) LPXLOPER12 "~wrapperName~"(LPXLOPER12 function(LPXLOPER12[] args) child)
	(
	";
	ret~="\t";
	foreach(i;0..numArgs)
	{
		ret~="LPXLOPER12 arg"~i.to!string;
		if (i<numArgs-1)
		{
			ret~=", ";
			if ((i%6==0) &&(i!=0))
				ret~="\n\t";
		}
	}
	ret~=")
	{
		LPXLOPER12[] args=[";
	foreach(i;0..numArgs)
	{
		ret~="arg"~i.to!string;
		if (i<numArgs-1)
			ret~=", ";
	}
	ret~="];
	size_t numArgs="~numArgs.to!string~";
	while(args[numArgs-1].xltype==xltypeMissing)
	{
		--numArgs;
		args.length=numArgs;
	}
	return (*child)(args);
}
";
	return ret;
}

mixin (makeMultiArgWrap("multiArgWrap30",30));
/+
multiArgWrap30(&simpleMulti);
+/
LPXLOPER12 simpleMulti(LPXLOPER12[] args)
{
	Variant[][][] niceArgs;
	bool success=true;
	foreach(arg;args)
	{
		auto conv=fromXLOPER12(arg);
		if (!conv.success)
		{
			success=false;
			break;
		}
		niceArgs~=conv.data;
	}
	return [args.length].makeXLOPER12;
}


