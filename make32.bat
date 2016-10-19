dmd -c -g -m32 -ofmyxll.obj example/myxll.d xlld/xll.d xlld/worksheet.d xlld/memorymanager.d xlld/memorypool.d xlld/xlcall.d xlld/xlcallcpp.d xlld/framework.d xlld/wrap.d
dmd -m32 -ofmyxll32.xll -L/IMPLIB myxll.obj generic32.def xlcall32d.lib -g -map
