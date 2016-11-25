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

import std.typecons: Flag, Yes;

enum StartingMemorySize = 10240;
enum MaxMemorySize=100*1024*1024;

private __gshared MemoryPool excelCallPool;


ubyte* GetTempMemory(Flag!"autoFree" autoFree = Yes.autoFree)(size_t numBytes)
{
    static if(autoFree) {
        // normally this would be done in a module constructor, but for
        // dmd-bug reasons the module constructor doesn't build
        // with either linker or compiler errors
        static bool init; // FIXME - see module constructor below
        if(!init) {
            excelCallPool = MemoryPool(StartingMemorySize);
            init = true;
        }
        return excelCallPool.allocate(numBytes).ptr;
    } else {
        import std.experimental.allocator;
        return theAllocator.allocate(numBytes).ptr;
    }
}

void FreeAllTempMemory() nothrow
{
    excelCallPool.freeAll;
}


struct MemoryPool {

    ubyte[] data;
    size_t curPos=0;


    this(size_t startingMemorySize) {
        import std.experimental.allocator;
        if (data.length==0)
            data=theAllocator.makeArray!(ubyte)(startingMemorySize);
        curPos=0;
    }

    ~this() {
        dispose;
    }

    void dispose() {
        import std.experimental.allocator;
        if(data.length>0)
            theAllocator.dispose(data);
        data.length=0;
        curPos=0;
    }

    ubyte[] allocate(size_t numBytes) {
        import std.experimental.allocator;
        import std.algorithm: min;

        if (numBytes<=0)
            return null;

        if (curPos + numBytes > data.length)
        {
            auto newAllocationSize = min(MaxMemorySize, data.length * 2);
            if (newAllocationSize <= data.length)
                return null;
            theAllocator.expandArray(data, newAllocationSize, 0);
        }

        auto lpMemory = data[curPos .. curPos+numBytes];
        curPos += numBytes;
        return lpMemory;
    }

    // Frees all the temporary memory by setting the index for available memory back to the beginning
    void freeAll() nothrow @nogc {
        curPos = 0;
    }
}
