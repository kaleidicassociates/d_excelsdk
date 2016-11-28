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
private __gshared bool gInit;


ubyte* GetTempMemory(Flag!"autoFree" autoFree = Yes.autoFree)(size_t numBytes)
{
    static if(autoFree) {
        // FIXME
        // normally this would be done in a module constructor, but for
        // dmd-bug reasons the module constructor doesn't build
        // with either linker or compiler errors
        if(!gInit) {
            excelCallPool = MemoryPool(StartingMemorySize);
            gInit = true;
        }
        return excelCallPool.allocate(numBytes).ptr;
    } else {
        import std.experimental.allocator;
        return theAllocator.allocate(numBytes).ptr;
    }
}

void FreeAllTempMemory() nothrow @nogc
{
    excelCallPool.freeAll;
}


struct MemoryPool {

    ubyte[] data;
    size_t curPos=0;


    this(size_t startingMemorySize) nothrow @nogc {
        import std.experimental.allocator: makeArray;

        if (data.length==0)
            data=allocator.makeArray!(ubyte)(startingMemorySize);
        curPos=0;
    }

    ~this() nothrow {
        dispose;
    }

    void dispose() nothrow {
        import std.experimental.allocator: dispose;

        if(data.length>0)
            allocator.dispose(data);
        data.length=0;
        curPos=0;
    }

    ubyte[] allocate(size_t numBytes) nothrow @nogc {
        import std.algorithm: min;
        import std.experimental.allocator: expandArray;

        if (numBytes<=0)
            return null;

        if (curPos + numBytes > data.length)
        {
            auto newAllocationSize = min(MaxMemorySize, data.length * 2);
            if (newAllocationSize <= data.length)
                return null;
            allocator.expandArray(data, newAllocationSize, 0);
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
