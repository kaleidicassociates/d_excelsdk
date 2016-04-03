module genXllDescr;
import xlltypes;

auto xllProcedure(string _xllText)() {
	return XllProcedure!_xllText.init;	
}

struct XllProcedure(wstring _xllText) {
	enum xllText = _xllText;
	alias xllText this;
	enum xllArgPosition = 0;
}

auto xllType(string _xllText)() {
	return XllType!_xllText.init;
}

struct XllType(wstring _xllText) {
	enum xllText = _xllText;
	alias xllText this;
	enum xllArgPosition = 1;
}

auto xllFunction(string _xllText)() {
	return XllFunction!_xllText.init;
}

struct XllFunction(wstring _xllText) {
	enum xllText = _xllText;
	alias xllText this;
	enum xllArgPosition = 2;
}

auto xllArgument(string _xllText)() {
	return XllArgument!_xllText.init;
}

struct XllArgument (wstring _xllText) {
	enum xllText = _xllText;
	alias xllText this;
	enum xllArgPosition = 3;
}

auto xllMacroType(string _xllText)() {
	return XllMacroType!_xllText.init;
}

struct XllMacroType (wstring _xllText) {
	enum xllText = _xllText;
	alias xllText this;
	enum xllArgPosition = 4;
}

auto xllCategory(string _xllText)() {
	return XllCategory!_xllText.init;
}

struct XllCategory (wstring _xllText) {
	enum xllText = _xllText;
	alias xllText this;
	enum xllArgPosition = 5;
}

auto xllShortcut(string _xllText)() {
	return XllShortcut!_xllText.init;
}

struct XllShortcut (wstring _xllText) {
	enum xllText = _xllText;
	alias xllText this;
	enum xllArgPosition = 6;
}

auto xllHelpTopic(string _xllText)() {
	return XllHelpTopic!_xllText.init;
}

struct XllHelpTopic (wstring _xllText) {
	enum xllText = _xllText;
	alias xllText this;
	enum xllArgPosition = 7;
}

auto xllFunctionHelp(wstring _xllText)() {
	return XllFunctionHelp!_xllText.init;
}

struct XllFunctionHelp (wstring _xllText) {
	enum xllText = _xllText;
	alias xllText this;
	enum xllArgPosition = 8;
}

auto xllArgumentHelp(string _xllText, uint argN = 0)() {
	return XllArgumentHelp!(_xllText.init, argN);
}

struct XllArgumentHelp(wstring _xllText, uint argN) {
	enum xllText = _xllText;
	alias xllText this;
	enum xllArgPosition = 9 + argN;
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
	wstring[10] args;

	this (T...) (T _args) /*if (args.allStatisfy!(t => (is(t.xllText : string) && is(t.xllArgPosition : uint)))) */ {
		import std.algorithm : startsWith, filter; 
		foreach(arg;_args) {
			static if (is(typeof(arg.xllArgPosition))) {
				pragma(msg, "Assigning arg (", arg.xllArgPosition, ") : ", arg.xllText);
				args[arg.xllArgPosition] = arg; 
			}
		}
	}
}



template descr(alias exportedFunction) {
	enum descr = mixin(descr_!exportedFunction);
	static assert(is(typeof(descr) == wstring[]));
}

char typeText(T)(T t) {
	// are C and F the same ?
	// F is modified in place
	// C is not
	static if (is(T == XlOper4)) {
		return "R";
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
}

wstring[10] descr_(alias Func)() {
	import std.traits : hasUDA, getUDAs;
	import std.algorithm : filter, startsWith;
	static assert(hasUDA!(exportedFunction, Xll));

	debug { uint argCtr; }
	auto xll = getUDAs!(Func, Xll)[0];
/*	pragma(msg, getUDAs!(Func, Xll)[0]);
	foreach(_member;__traits(derivedMembers, typeof(xll))) {
		static if (_member.startsWith("Xll")) {
 		auto member = __traits(getMember, xll, _member);
			debug { assert(typeof(member).xllArgPosition == argCtr++); }
			result[typeof(member).xllArgPosition] = member;
		}

	}
*/

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

extern (C) @Xll(xllProcedure!("AddOne"), xllCategory!("SimpleMath"), xllType!("BB"), xllFunctionHelp!("Adds one to the argument")) 
	double exportedFunction(double ctr) { return ctr++; }

static assert(descr_!(exportedFunction) == ["AddOne"w, "BB"w, ""w, ""w, ""w, "SimpleMath"w, ""w, ""w, "Adds one to the argument"w, ""w ]);
