// Developer test host only; production loads the XLL directly inside WPS.
#define UNICODE
#define _UNICODE
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
typedef int (__cdecl *CALL)(void);
int wmain(int argc,WCHAR **argv) {
  if(argc!=2&&argc!=3&&argc!=5&&argc!=6)return 2;
  HMODULE dll=LoadLibraryW(argv[1]);if(!dll){printf("load_error=%lu\n",GetLastError());return 3;}
  CALL start=(CALL)GetProcAddress(dll,"OfficeBridgeProbeTestStart"),state=(CALL)GetProcAddress(dll,"OfficeBridgeProbeTestState"),stop=(CALL)GetProcAddress(dll,"xlAutoRemove");
  if(!start||!state||!stop){FreeLibrary(dll);return 4;}
  int started=0;
  if(argc==6) {
    typedef int (__cdecl *RELAY)(DWORD,DWORD,DWORD);
    RELAY relay=(RELAY)GetProcAddress(dll,"OfficeBridgeProbeTestStartRelay");
    if(!relay||wcscmp(argv[2],L"--relay")){FreeLibrary(dll);return 4;}
    started=relay((DWORD)wcstoul(argv[3],0,10),(DWORD)wcstoul(argv[4],0,10),(DWORD)wcstoul(argv[5],0,10));
  } else if(argc==5) {
    typedef int (__cdecl *FULL)(DWORD,DWORD);
    FULL full=(FULL)GetProcAddress(dll,"OfficeBridgeNativeTestStart");
    if(!full||wcscmp(argv[2],L"--full")){FreeLibrary(dll);return 4;}
    started=full((DWORD)wcstoul(argv[3],0,10),(DWORD)wcstoul(argv[4],0,10));
  } else if(argc==3) {
    typedef int (__cdecl *PORT)(DWORD);
    PORT port=(PORT)GetProcAddress(dll,"OfficeBridgeProbeTestStartPort");
    if(!port){FreeLibrary(dll);return 4;}started=port((DWORD)wcstoul(argv[2],0,10));
  } else started=start();
  if(!started){FreeLibrary(dll);return 5;}
  ULONGLONG deadline=GetTickCount64()+30000;
  while(state()&&GetTickCount64()<deadline)Sleep(50);
  if(state()){stop();for(int n=0;n<100&&state();n++)Sleep(50);}
  int pending=state();FreeLibrary(dll);return pending?6:0;
}