/**
	Ported from Microsoft's Excel SDK MemoryManager.cpp by Laeeth Isharc.  See Excel SDK for copyright

	The memory manager class is an update to the memory manager
	in the previous release of the framework.  This class provides
	each thread with an array of bytes to use as temporary memory.
	The size of the array, and the methods for dealing with the
	memory explicitly, is in the class MemoryPool.

	MemoryManager handles assigning of threads to pools, and the
	creation of new pools when a thread asks for memory the first
	time.  Using a singleton class, the manager provides an interface
	to C code into the manager.  The number of unique pools starts
	as TotalMemoryPools, defined in MemoryManager.h.  When a new thread
	needs a pool, and the current set of pools are all assigned,
	the number of pools increases by a factor of two.
*/

module xlld.memorymanager;
import xlld.xlcall;
import xlld.xlcallcpp;
import std.experimental.allocator;
import core.sys.windows.windows;

//
// Total amount of memory to allocate for all temporary XLOPERs
//

enum MEMORYSIZE=10240;
enum MaxMemorySize=100*1024*1024;

MemoryPool excelCallPool;

static this()
{
	excelCallPool.start();
}

static ~this()
{
	excelCallPool.dispose();
}

struct MemoryPool
{
	DWORD m_dwOwner=cast(DWORD)-1;			// ID of ownning thread
	ubyte[] data;
	size_t curPos=0;
	ubyte[MEMORYSIZE] m_rgchMemBlock;		// Memory for temporary XLOPERs
	size_t m_ichOffsetMemBlock=0;	// Offset of next memory block to allocate


	void start()
	{
		if (data.length==0)
			data=theAllocator.makeArray!(ubyte)(MEMORYSIZE);
		curPos=0;
	}
	// An empty destructor - see reasoning below
	//
	void dispose()
	{
		if(data.length>0)
			theAllocator.dispose(data);
		data.length=0;
		curPos=0;
	}

	ubyte[] allocate(size_t numBytes)
	{
            import std.algorithm: min;
		ubyte[] lpMemory;

		if (numBytes<=0)
			return null;

		if (curPos + numBytes > data.length)
		{
			auto newAllocationSize=min(MaxMemorySize,data.length*2);
			if (newAllocationSize<=data.length)
				return null;
			theAllocator.expandArray(data,newAllocationSize,0);
		}

		lpMemory = data[curPos..curPos+numBytes];
		curPos+=numBytes;
		return lpMemory;
	}

	// Frees all the temporary memory by setting the index for available memory back to the beginning

	void freeAll() nothrow @nogc
	{
		curPos = 0;
	}
}



// void WINAPI xlAutoFree12(LPXLOPER12 arg)
// {
//     if (arg.xltype & xlbitDLLFree)
//  		freeXLOper(arg);
// }
