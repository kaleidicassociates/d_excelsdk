module genXllDescr;

struct XllProcedure {
	wstring xllText;
	alias xllText this;
	enum xllArgPosition = 0;
}

struct XllType {
	wstring xllText;
	alias xllText this;
	enum xllArgPosition = 1;
}
struct XllFunction {
	wstring xllText;
	alias xllText this;
	enum xllArgPosition = 2;
}
struct XllArgument {
	wstring xllText;
	alias xllText this;
	enum xllArgPosition = 3;
}
struct XllMacroType {
	wstring xllText;
	alias xllText this;
	enum xllArgPosition = 4;
}
struct XllCategory {
	wstring xllText;
	alias xllText this;
	enum xllArgPosition = 5;
}
struct XllShortcut {
	wstring xllText;
	alias xllText this;
	enum xllArgPosition = 6;
}
struct XllHelpTopic {
	wstring xllText;
	alias xllText this;
	enum xllArgPosition = 7;
}
struct XllFunctionHelp {
	wstring xllText;
	alias xllText this;
	enum xllArgPosition = 8;
}

struct XllArgumentHelp {
	wstring xllText;
	alias xllText this;
	enum xllArgPosition = 9;
}

struct Xll {
	XllProcedure xllProcedure;
	XllType xllType;
	XllFunction xllFunction;
	XllArgument xllArgument;
	XllMacroType xllMacroType;
	XllCategory xllCategory;
	XllShortcut xllShortcut;
	XllFunctionHelp xllFunctionHelp;
	XllArgumentHelp xllArgumentHelp;

	this (T...) (T args) if (allStatisfy!(t => (is(t.xllText == wstring) && is(t.xllArgPosition : uint)))) {
		import std.algorithm : startsWith, filter; 
		foreach(arg;args) {
			foreach(member;__traits(derivedMembers, this).filter!(a => a.startWith("Xll"))) {
				static if (is(typeof(member) == typeof(arg))) {
					assert(__traits(getMember, this, member) == member.init);
					__traits(getMember, this, member) = arg;
				}
			}
		}
	}
}



import xlltypes;
template descr(alias exportedFunction) {
	enum descr = mixin(descr_!exportedFunction);
	static assert(is(typeof(descr) == wstring[]));
}

wchar typeText(T)(T t) {
	// are C and F the same ?
	// F is modified in place
	// C is not
	static if (is(T == LxOper)) {

	} else static if (is(T == double)) {
		return "B"w;
	} else static if (is(T == double*)) {
		return "E"w;
	} else static if (is(T == Boolean)) {
		return "A"w;
	} else static if (is(T == Boolean*)) {
		return "L"w;
	} else static if (is(T == const char*)) {
		return "C"w;
	} else static if (is(T == char*)) {
		return "F"w;
	}
}

wstring[10] descr_(alias exportedFunction)() {
	import std.traits : hasUDA;
	import std.algorithm : filter;
	static assert(hasUDA!(exportedFunction, Xll));
	wstring[10] result;

	debug { uint argCtr; }
	auto xll = getUDA!(exportedFunction, Xll);
	
	foreach(member;__traits(derivedMembers, xll).filter!(is(member.xllArgPosition))) {
		debug { assert(member.xllArgPosition == argCtr++); }
		result[member.xllArgPosition] = member;
	}

	return result;


	/* Output looks like :
	[ "Func1"w,                                     // Procedure1
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
