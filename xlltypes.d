module xlltypes;
import std.exception : enforce;
import memorymanager : GetTempMemory;


struct Boolean {
	short _value;
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
	xltypeNum = 0x0001,
	xltypeStr = 0x0002,
	xltypeBool= 0x0004,
	xltypeRef =  0x0008,
	xltypeErr = 0x0010,
	xltypeFlow = 0x0020,
	xltypeMulti = 0x0040,
	xltypeMissing = 0x0080,
	xltypeNil =  0x0100,
	xltypeSRef = 0x0400,
	xltypeInt = 0x0800,
	
	xltypeBigData = (xltypeStr | xltypeInt)
}

struct Xloper12
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
		ushort _bool; /* xltypeBool */
		// (ENDXXX)
		
		ushort error; /* xltypeErr */
		short integer; /* xltypeInt */
	//	Sref sref;
		
	}
	
	XloperType type;
	
	this(string _string) {
		this.type = XloperType.xltypeStr;
		
	}

}


