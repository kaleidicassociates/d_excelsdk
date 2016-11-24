/**
	MemoryManager.D

	Ported from MemoryManager.cpp by Laeeth Isharc


//
// Purpose:     The memory manager class is an update to the memory manager
//              in the previous release of the framework.  This class provides
//              each thread with an array of bytes to use as temporary memory.
//              The size of the array, and the methods for dealing with the
//              memory explicitly, is in the class MemoryPool.
//
//              MemoryManager handles assigning of threads to pools, and the
//              creation of new pools when a thread asks for memory the first
//              time.  Using a singleton class, the manager provides an interface
//              to C code into the manager.  The number of unique pools starts
//              as MEMORYPOOLS, defined in MemoryManager.h.  When a new thread
//              needs a pool, and the current set of pools are all assigned,
//              the number of pools increases by a factor of two.
//
// Platform:    Microsoft Windows
//
///***************************************************************************
*/
module xlld.memorymanager;

private __gshared MemoryManager vpmm;



import core.sys.windows.windows;
version (Windows) {
	@nogc nothrow extern(Windows) uint GetCurrentThreadId();
} else {
	uint GetCurrentThreadId() {
		import core.thread;
		return cast(uint)Thread.getThis().id;
	}
}
import xlld.memorypool;

//
// Total number of memory allocation pools to manage
//
enum MEMORYPOOLS=4;

struct MemoryManager
{
	private int _numPools = 0;
	private int _maxNumPools = MEMORYPOOLS;
	static private MemoryPool[] _pools = new MemoryPool[MEMORYPOOLS]; // Storage for the memory pools

	//
	// Returns the singleton class, or creates one if it doesn't exit
	//
	static MemoryManager* GetManager() nothrow @nogc
	{
            return &vpmm;
        }

	//
	// Method that will query the correct memory pool of the calling
	// thread for a set number of bytes.  Returns 0 if there was a
	// failure in getting the memory.
	//
	ubyte* CPP_GetTempMemory(size_t cByte)
	{
		auto pmp = GetMemoryPool(GetCurrentThreadId());
                return pmp ? pmp.GetTempMemory(cByte) : null;
	}

	//
	// Method that tells the pool owned by the calling thread that
	// it is free to reuse all of its memory
	//
	void CPP_FreeAllTempMemory() nothrow
	{
		auto pmp = GetMemoryPool(GetCurrentThreadId());
                return pmp ? pmp.FreeAllTempMemory : null;
	}

	//
	// Method iterates through the memory pools in an attempt to find
	// the pool that matches the given thread ID. If a pool is not found,
	// it creates a new one
	//
	private MemoryPool* GetMemoryPool(uint dwThreadID) nothrow
	{
            import std.algorithm: find;
            import std.array: empty, front;

            auto pool = _pools.find!(a => a.m_dwOwner == dwThreadID);
            return pool.empty
                ? CreateNewPool(dwThreadID) //didn't find the owner, make a new one
                : &pool.front;
	}

	//
	// Will assign an unused pool to a thread; should all pools be assigned,
	// it will grow the number of pools available.
	//
	private MemoryPool* CreateNewPool(uint dwThreadID) nothrow
	{
		if (_numPools >= _maxNumPools)
		{
			GrowPools();
		}
		_pools[_numPools++].m_dwOwner = dwThreadID;

		return &_pools[_numPools - 1];
	}

	//
	// Increases the number of available pools by a factor of two. All of
	// the old pools have their memory pointed to by the new pools. The
	// memory for the new pools that get replaced is first freed. The reason
	// ~MemoryPool() can't free its array is in this method - they would be
	// deleted when the old array of pools is freed at the end of the method,
	// despite the fact they are now being pointed to by the new pools.
	//
	void GrowPools() nothrow
	{
		auto impMaxNew = 2 * _maxNumPools;
		_pools.length = 2 * _maxNumPools;
		_maxNumPools = impMaxNew;
	}
}

extern (C++) :
//
// Interface for C callers to ask for memory
ubyte* MGetTempMemory(size_t cByte)
{
	return MemoryManager.GetManager().CPP_GetTempMemory(cByte);
}

//
// Interface for C callers to allow their memory to be reused
void MFreeAllTempMemory() nothrow
{
	MemoryManager.GetManager().CPP_FreeAllTempMemory();
}

void* GetTempMemory(size_t size) {
	return MGetTempMemory(size);
}
