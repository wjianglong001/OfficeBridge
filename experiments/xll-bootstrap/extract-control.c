// Isolates resource extraction from WPS. Never launches the embedded backend or
// writes into the XLL's actual startup cache, so it cannot preinstall a payload.
#define XLL_PROBE_EXTRACT_ONLY
#include "bootstrap.c"
int wmain(int argc,WCHAR **argv) {
  if(argc!=2)return 2;
  module=LoadLibraryExW(argv[1],0,LOAD_LIBRARY_AS_DATAFILE);
  if(!module)return 3;
  DWORD result=start(0);
  FreeLibrary(module);
  return (int)result;
}
