module genxlldescr;
import xlltypes;
const pure {
auto xllProcedure(string _xllText)() {
	return XllProcedure!_xllText.init;	
}

uint xllProcedure() {
	return 0;
}

struct XllProcedure(string _xllText) {
	enum xllText = _xllText;
	alias xllText this;
	enum xllArgPosition = xllProcedure;
}

auto xllType(string _xllText)() {
	return XllType!_xllText.init;
}

uint xllType() {
	return 1;
}

struct XllType(string _xllText) {
	enum xllText = _xllText;
	alias xllText this;
	enum xllArgPosition = xllType;
}

auto xllFunction(string _xllText)() {
	return XllFunction!_xllText.init;
}

uint xllFunction() {
	return 2;
}

struct XllFunction(string _xllText) {
	enum xllText = _xllText;
	alias xllText this;
	enum xllArgPosition = xllFunction;
}

auto xllArgument(string _xllText)() {
	return XllArgument!_xllText.init;
}

struct XllArgument (string _xllText) {
	enum xllText = _xllText;
	alias xllText this;
	enum xllArgPosition = 3;
}

auto xllMacroType(string _xllText)() {
	return XllMacroType!_xllText.init;
}

struct XllMacroType (string _xllText) {
	enum xllText = _xllText;
	alias xllText this;
	enum xllArgPosition = 4;
}

auto xllCategory(string _xllText)() {
	return XllCategory!_xllText.init;
}

struct XllCategory (string _xllText) {
	enum xllText = _xllText;
	alias xllText this;
	enum xllArgPosition = 5;
}

auto xllShortcut(string _xllText)() {
	return XllShortcut!_xllText.init;
}

struct XllShortcut (string _xllText) {
	enum xllText = _xllText;
	alias xllText this;
	enum xllArgPosition = 6;
}

auto xllHelpTopic(string _xllText)() {
	return XllHelpTopic!_xllText.init;
}

struct XllHelpTopic (string _xllText) {
	enum xllText = _xllText;
	alias xllText this;
	enum xllArgPosition = 7;
}

auto xllFunctionHelp(string _xllText)() {
	return XllFunctionHelp!_xllText.init;
}

struct XllFunctionHelp (string _xllText) {
	enum xllText = _xllText;
	alias xllText this;
	enum xllArgPosition = 8;

}

auto xllArgumentHelp(string _xllText, uint argN = 0)() {
	return XllArgumentHelp!(_xllText.init, argN);
}

struct XllArgumentHelp(string _xllText, uint argN) {
	enum xllText = _xllText;
	alias xllText this;
	enum xllArgPosition = 9 + argN;
}
}
struct Xll {
/*
	XllProcedure xllProcedure;
	XllType xllType;
	XllFunction xllFunction;
	XllArgument xllArgument;
	XllMacroType xllMacroType;
	XllCategory xllCategory;
	XllShortcut xllShortcut;
	XllFunctionHelp xllFunctionHelp;
	XllArgumentHelp xllArgumentHelp;
*/
	string[10] args;

	this (T...) (T _args) /*if (args.allStatisfy!(t => (is(t.xllText : string) && is(t.xllArgPosition : uint)))) */ {
		foreach(arg;_args) {
			static if (is(typeof(arg.xllArgPosition))) {
				args[arg.xllArgPosition] = arg; 
			}
		}
	}
}



template descr(alias exportedFunction) {
	enum descr = mixin(descr_!exportedFunction);
	static assert(is(typeof(descr) : string[]));
}

string typeText(T)(T t) {
	// are C and F the same ?
	// F is modified in place
	// C is not
	static if (is(T == Xloper12)) {
		return "O";
	} else static if (is(T == double)) {
		return "B";
	} else static if (is(T == double*)) {
		return "E";
	} else static if (is(T == short)) {
		return "I";
	} else static if (is(T == short*)) {
		return "M";
	} else static if (is(T == int)) {
		return "J";
	} else static if (is(T == int*)) {
		return "N";
	} else static if (is(T == Boolean)) {
		return "A";
	} else static if (is(T == Boolean*)) {
		return "L";
	} else static if (is(T == const char*)) {
		return "C";
	} else static if (is(T == char*)) {
		return "F";
	}
	
	assert(0, "Cannot find mangle for "  ~ T.stringof);
}

string[10] descr_(alias Func)() {
	import std.traits : hasUDA, getUDAs;
	import std.algorithm : filter, startsWith;
	static assert(hasUDA!(exportedFunction, Xll));
	
	static if (!getUDAs!(Func, Xll).length) {
		auto xll = Xll.init;
	} else {
		auto xll = getUDAs!(Func, Xll)[0];
	}
 	
/*	pragma(msg, getUDAs!(Func, Xll)[0]);
	foreach(_member;__traits(derivedMembers, typeof(xll))) {
		static if (_member.startsWith("Xll")) {
 		auto member = __traits(getMember, xll, _member);
			debug { assert(typeof(member).xllArgPosition == argCtr++); }
			result[typeof(member).xllArgPosition] = member;
		}

	}
*/

	if(!xll.args[xllProcedure]) {
		xll.args[xllProcedure] = __traits(identifier, Func);
	}
	if(xll.args[xllType] == "") /*XllType*/ {
		import std.range : join, iota;
		import std.array : array;
		import std.algorithm : map;
		import std.traits : Parameters, ReturnType;
		static if (is(ReturnType!Func)) {
			xll.args[xllType] = typeText(ReturnType!Func.init);
		} else {
			static assert(0, "Functions without returnType are unsupported");
		}
		foreach(p;Parameters!Func) {
			xll.args[xllType] ~= typeText(p.init);
		}

	}
	if(xll.args[xllFunction] == "") {
		xll.args[xllFunction] = __traits(identifier, Func);
	}
	//pragma(msg, "mangle for function ", __traits(identifier, Func), "is ", xll.args[xllType]);
	return xll.args;


	/* Output looks like :
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
	]
*/
	 
}


//// unittest

extern (C) @Xll(xllCategory!("SimpleMath"), xllFunctionHelp!("Adds one to the argument")) 
	double exportedFunction(double ctr) { return ctr++; }
//pragma(msg, descr_!exportedFunction);
static assert(descr_!(exportedFunction) == ["exportedFunction", "BB", "exportedFunction", "", "", "SimpleMath", "", "", "Adds one to the argument", ""]);
