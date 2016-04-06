module xlltypes;
import std.exception : enforce;
import memorymanager : GetTempMemory;


struct Boolean {
	@property short _value;
	@property bool asBool() pure {
		return _value != 0;
	}

	@property void asBool(bool val) pure {
		val ? _value = 1 : _value = 0;
	}

	alias this asBool;
}

struct CountedString4 {
	private ubyte* _data;

	@property string toString() {
		return cast(string) _data[1 .. length + 1];
	} 

	@property ubyte length() {
		return _data[0];
	}

	this(ubyte* data) {
		this._data = data;
	}

/*	this(string _string) {
		import core.stdc.string : memcpy;
		enforce(_string.length < 255, "Strings longer then 255 chars are not supported");
		this._data = GetTempMemory(XLOPER12.sizeof + (_string.length+1));
		memcpy(_data + 1, _string.ptr, _string.length);
		this._data[0] = cast(ubyte) _string.length;
	}
*/	
	alias toString this;
}



struct XlSRef {
	ushort rwFirst;
	ushort rwLast;
	ubyte colFirst;
	ubyte colLast;
}

enum XloperType : ushort {
	Num = 0x0001,
	Str = 0x0002,
	Bool = 0x0004,
	Ref =  0x0008,
	Err = 0x0010,
	Flow = 0x0020,
	Multi = 0x0040,
	Missing = 0x0080,
	Nil =  0x0100,
	SRef = 0x0400,
	Int = 0x0800,
	
	BigData = (Str | Int)
}



struct Xloper12 //XXX Xloper12.sizeof has to be 32
{
/*
Blocked by DMD bug
	static struct Sref {
		ushort count;
		XlSRef reference; 
	}
*/
	union {
		double num; /* xltypeNum */
		char* _string; /* xltypeStr */
		
		// (XXX): really ??
		Boolean _bool; /* xltypeBool */
		// (ENDXXX)
	
		ushort error; /* xltypeErr */
		short integer; /* xltypeInt */
	//	Sref sref;
	
		static struct Xloper12Array {
			private static struct autoConvertSlice(T) {
				Xloper12* parray;
				static if (isArray!T) {
					uint rows;
					uint columns;
					
					this(Xloper12Array array) {
						parray = array.parray;
						rows = array.rows;
						columns = array.columns;
					}
					
					ElementType!T opIndex(uint x, uint y) {
						assert(x<columns);
						assert(y<rows);
				
						auto ret = *(parray + (y*columns) + x);
						return ret!(v => cast(T) v);		
					}
					
				} else {
					ulong length;
					
					this(Xloper12Array array) {
						parray = array.parray;
						length = array.columns * array.rows;
					}
					
					T opIndex(ulong pos) {
						auto ret = *(parray + pos);
						return ret!(v => cast(T) v);
					}
					
					@property T[] toArray() {
						import typeutils.heaparray;
						import std.algorithm : map;
						import std.array.array;
						return TempHeapSlice!(TempAllocateExcel, T)(this.toArrayRange.map!(v => v!(v2 => cast(T) v2)).array);
					}
				}
			}
		
		
			Xloper12* parray;
			uint rows;
			uint columns;
			
			Xloper12* opIndex(uint x, uint y) {
				assert(x<columns);
				assert(y<rows);
				
				return parray + (y*columns) + x;
			}
			
			auto toSlice (T)() {
				static if (isArray!T) { 
					static if (!isArray!(ElementTypes!T)) {
						//2d array.
						return autoConvertSlice!T(this);
					} else {
						// it's 3d or more
						assert(0, "only 1d or 2d arrays are supported");
					}
				} else {
					//1d slice
					return autoConvertSlice!T(this);
				}
			}
		}
		
		Xloper12Array array;
		
		static struct Xloper12Flow {
			union {
				int level;			/* xlflowRestart */
				int tbctrl;			/* xlflowPause */
				uint* idSheet;		/* xlflowGoto */
			}
			uint rw;				       	/* xlflowGoto */
			uint col;			       	/* xlflowGoto */
			ubyte xlflow;
		}			
		
		Xloper12Flow flow;
	}

	XloperType type;
	
	this(string _string) {
		type = XloperType.Str;
		assert(0, "strings are not supported at the moment");
	}
	
	this(short val) {
		type = XloperType.Int;
		integer = val;
	}
	
	this(double val) {
		type = XloperType.Num;
		num = val;
	}
	
	this (bool val) {
		type = XloperType.Bool;
	}
	
	R opCall(alias handler)() {
		switch(type) with (XloperType) {
			case Int : 
				static if (is(typeof(handler(integer.init)))) {
					return handler(integer);
				} else {
					assert(0, "Handler is not callable with type short");
				}
			case Num : 
				static if (is(typeof(handler(num.init)))) {
					return handler(num);
				} else {
					assert(0, "Hander is not callable with type double");
				}
			case Bool : 
				static if (is(typeof(handler(_bool.init)))) {
					return handler(_bool);
				} else {
					assert(0, "Hander is not callable with type (Boolean / bool)");
				}
			
			default :
				import std.conv : to;
				assert(0, "Type " ~ to!string(type) ~ " is not supported right now");
		}
	}
}


static assert(Xloper12.Xloper12Flow.sizeof == 16);
