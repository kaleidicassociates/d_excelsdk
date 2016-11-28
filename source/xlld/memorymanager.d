/**
	MemoryManager.D

	Ported from MemoryManager.cpp by Laeeth Isharc
//
// Platform:    Microsoft Windows
//
///***************************************************************************
*/
module xlld.memorymanager;

import std.typecons: Flag, Yes;
import std.experimental.allocator.mallocator: Mallocator;

alias Allocator = Mallocator;
alias allocator = Allocator.instance;

enum StartingMemorySize = 10240;
enum MaxMemorySize=100*1024*1024;

private __gshared MemoryPool excelCallPool;


ubyte* GetTempMemory(Flag!"autoFree" autoFree = Yes.autoFree)(size_t numBytes)
{
    static if(autoFree) {
        // FIXME
        // normally this would be done in a module constructor, but for
        // dmd-bug reasons the module constructor doesn't build
        // with either linker or compiler errors
        static bool init;
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
