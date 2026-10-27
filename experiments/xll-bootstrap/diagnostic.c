#define UNICODE
#define _UNICODE
#include <windows.h>
#include <stdio.h>
#include <wchar.h>

typedef struct { ULONG attributes; ACCESS_MASK granted; ULONG handles; ULONG pointers; ULONG reserved[10]; } OBJECT_BASIC;
typedef LONG (NTAPI *QUERY_OBJECT)(HANDLE,ULONG,PVOID,ULONG,PULONG);
static HMODULE image;
static volatile LONG running;
static WCHAR results[2048];
static void attempt(const WCHAR *dir,const WCHAR *name,const void *bytes,DWORD length) {
  WCHAR path[32768],line[256];
  swprintf(path,32768,L"%ls\\%ls",dir,name);
  SetLastError(0);
  HANDLE file=CreateFileW(path,GENERIC_WRITE,0,0,CREATE_NEW,FILE_ATTRIBUTE_NORMAL,0);
  DWORD createError=GetLastError(),written=0,error=0,type=0;
  OBJECT_BASIC info={0};
  QUERY_OBJECT query=(QUERY_OBJECT)GetProcAddress(GetModuleHandleW(L"ntdll.dll"),"NtQueryObject");
  LONG queryStatus=-1;BOOL ok=FALSE;
  if(file!=INVALID_HANDLE_VALUE) {
    type=GetFileType(file);
    if(query)queryStatus=query(file,0,&info,sizeof(info),0);
    SetLastError(0);
    ok=WriteFile(file,bytes,length,&written,0);error=GetLastError();
    CloseHandle(file);
  }
  swprintf(line,256,L"%ls: create=%lu type=%lu access=%08lx query=%08lx write=%d bytes=%lu/%lu error=%lu\n",name,createError,type,info.granted,queryStatus,ok,written,length,error);
  wcscat(results,line);
}
static DWORD WINAPI diagnose(void *arg) {
  (void)arg;
  WCHAR dir[32768],suffix[100],line[256];
  DWORD n=GetEnvironmentVariableW(L"LOCALAPPDATA",dir,30000);
  if(!n||n>29000)return 1;
  swprintf(suffix,100,L"\\OfficeBridge\\XllDiagnostic-%lu-%llu",GetCurrentProcessId(),GetTickCount64());
  wcscat(dir,suffix);
  if(!CreateDirectoryW(dir,0))return 2;
  HANDLE token=0;DWORD size=0,appContainer=0;TOKEN_ELEVATION elevation={0};
  if(OpenProcessToken(GetCurrentProcess(),TOKEN_QUERY,&token)) {
    GetTokenInformation(token,TokenIsAppContainer,&appContainer,sizeof(appContainer),&size);
    GetTokenInformation(token,TokenElevation,&elevation,sizeof(elevation),&size);
    swprintf(line,256,L"pid=%lu restricted=%d container=%lu elevated=%lu\n",GetCurrentProcessId(),IsTokenRestricted(token),appContainer,elevation.TokenIsElevated);
    wcscat(results,line);CloseHandle(token);
  }
  const char plain[]="OfficeBridge diagnostic synthetic text";
  attempt(dir,L"plain.txt",plain,sizeof(plain)-1);
  HRSRC res=FindResourceW(image,MAKEINTRESOURCEW(101),RT_RCDATA);
  HGLOBAL loaded=res?LoadResource(image,res):0;
  const void *bytes=loaded?LockResource(loaded):0;
  DWORD length=res?SizeofResource(image,res):0;
  swprintf(line,256,L"resource=%lu valid=%d\n",length,bytes!=0);wcscat(results,line);
  if(bytes&&length>=32768)attempt(dir,L"payload.tmp",bytes,32768);
  MessageBoxW(0,results,L"OfficeBridge XLL Diagnostic",MB_OK|MB_ICONINFORMATION);
  return 0;
}
__declspec(dllexport) int __cdecl xlAutoOpen(void) {
  if(InterlockedCompareExchange(&running,1,0))return 1;
  GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS|GET_MODULE_HANDLE_EX_FLAG_PIN,(LPCWSTR)&xlAutoOpen,&image);
  HANDLE thread=CreateThread(0,0,diagnose,0,0,0);if(thread)CloseHandle(thread);
  return 1;
}
__declspec(dllexport) int __cdecl xlAutoAdd(void){return 1;}
__declspec(dllexport) int __cdecl xlAutoClose(void){return 1;}
__declspec(dllexport) int __cdecl xlAutoRemove(void){return 1;}
