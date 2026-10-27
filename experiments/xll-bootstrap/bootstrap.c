#define UNICODE
#define _UNICODE
#include <windows.h>
#include <winhttp.h>
#include <wincrypt.h>
#include <stdio.h>
#include <string.h>
#include <wchar.h>
#include "build-config.h"

static HMODULE module;
static volatile LONG running;
static BOOL health(void) {
  BOOL ok=FALSE;
  HINTERNET session=WinHttpOpen(L"OfficeBridge-XLL-Probe/1",WINHTTP_ACCESS_TYPE_NO_PROXY,0,0,0);
  if(!session)return FALSE;
  WinHttpSetTimeouts(session,1000,1000,1000,1000);
  HINTERNET connection=WinHttpConnect(session,L"127.0.0.1",43191,0);
  HINTERNET request=connection?WinHttpOpenRequest(connection,L"GET",L"/probe/health",0,0,0,0):0;
  if(request && WinHttpSendRequest(request,L"X-OfficeBridge-Xll-Probe: " XLL_TOKEN L"\r\n",-1,0,0,0,0) && WinHttpReceiveResponse(request,0)) {
    DWORD code=0,size=sizeof(code);
    if(WinHttpQueryHeaders(request,WINHTTP_QUERY_STATUS_CODE|WINHTTP_QUERY_FLAG_NUMBER,0,&code,&size,0) && code==200)ok=TRUE;
  }
  if(request)WinHttpCloseHandle(request);
  if(connection)WinHttpCloseHandle(connection);
  WinHttpCloseHandle(session);
  return ok;
}
static BOOL correct_file(const WCHAR *path) {
  HANDLE file=CreateFileW(path,GENERIC_READ,FILE_SHARE_READ,0,OPEN_EXISTING,FILE_FLAG_SEQUENTIAL_SCAN,0);
  if(file==INVALID_HANDLE_VALUE)return FALSE;
  HCRYPTPROV provider=0; HCRYPTHASH hash=0; BOOL ok=FALSE;
  if(CryptAcquireContextW(&provider,0,0,PROV_RSA_AES,CRYPT_VERIFYCONTEXT) && CryptCreateHash(provider,CALG_SHA_256,0,0,&hash)) {
    BYTE bytes[32768],digest[32]; DWORD count=0,length=32; BOOL readOK=TRUE;
    for(;;) { if(!ReadFile(file,bytes,sizeof(bytes),&count,0)){readOK=FALSE;break;} if(!count)break; if(!CryptHashData(hash,bytes,count,0)){readOK=FALSE;break;} }
    if(readOK && CryptGetHashParam(hash,HP_HASHVAL,digest,&length,0) && length==32)ok=memcmp(digest,BACKEND_SHA256,32)==0;
  }
  if(hash)CryptDestroyHash(hash);
  if(provider)CryptReleaseContext(provider,0);
  CloseHandle(file);return ok;
}
static BOOL ensure_directory(const WCHAR *path) {
  if(!CreateDirectoryW(path,0) && GetLastError()!=ERROR_ALREADY_EXISTS)return FALSE;
  DWORD attributes=GetFileAttributesW(path);
  return attributes!=INVALID_FILE_ATTRIBUTES && (attributes&FILE_ATTRIBUTE_DIRECTORY) && !(attributes&FILE_ATTRIBUTE_REPARSE_POINT);
}
static DWORD WINAPI start(void *unused) {
  (void)unused;
  const WCHAR *stage=L"directory";
  WCHAR path[32768],exe[32768],temp[32768];
  DWORD n=GetEnvironmentVariableW(L"LOCALAPPDATA",path,32000);
  if(!n || n>30000)goto failed;
  wcscat(path,L"\\OfficeBridge");if(!ensure_directory(path))goto failed;
#ifdef XLL_PROBE_EXTRACT_ONLY
  wcscat(path,L"\\XllProbeControl");if(!ensure_directory(path))goto failed;
#else
  wcscat(path,L"\\XllProbe");if(!ensure_directory(path))goto failed;
#endif
  wcscat(path,L"\\" BUILD_ID);if(!ensure_directory(path))goto failed;
  swprintf(exe,32768,L"%ls\\OfficeBridgeProbe.exe",path);
  swprintf(temp,32768,L"%ls\\OfficeBridgeProbe-%lu-%llu.tmp",path,GetCurrentProcessId(),GetTickCount64());
  HANDLE mutex=CreateMutexW(0,FALSE,L"Local\\OfficeBridge.XllProbe." BUILD_ID);
  if(!mutex || WaitForSingleObject(mutex,15000)!=WAIT_OBJECT_0)goto failed;
  if(!correct_file(exe)) {
    stage=L"resource";
    HRSRC resource=FindResourceW(module,MAKEINTRESOURCEW(101),RT_RCDATA);
    HGLOBAL loaded=resource?LoadResource(module,resource):0;
    DWORD length=resource?SizeofResource(module,resource):0;
    void *data=loaded?LockResource(loaded):0;
    if(!data || !length){ReleaseMutex(mutex);CloseHandle(mutex);goto failed;}
    stage=L"extract";
    HANDLE file=CreateFileW(temp,GENERIC_WRITE,0,0,CREATE_NEW,FILE_ATTRIBUTE_NORMAL,0);
    stage=L"create_file";
    if(file==INVALID_HANDLE_VALUE){ReleaseMutex(mutex);CloseHandle(mutex);goto failed;}
    DWORD written=0,offset=0;
    BOOL good=TRUE;
    stage=L"write_resource";
    while(offset<length){DWORD block=length-offset;if(block>32768)block=32768;if(!WriteFile(file,(BYTE*)data+offset,block,&written,0)||written!=block){good=FALSE;break;}offset+=written;}
    DWORD writeError=GetLastError();
    if(file!=INVALID_HANDLE_VALUE)CloseHandle(file);
    if(!good){SetLastError(writeError);ReleaseMutex(mutex);CloseHandle(mutex);goto failed;}
    stage=L"verify_resource";
    if(!correct_file(temp)){ReleaseMutex(mutex);CloseHandle(mutex);goto failed;}
    stage=L"rename_resource";
    if(!MoveFileExW(temp,exe,MOVEFILE_REPLACE_EXISTING|MOVEFILE_WRITE_THROUGH)) { ReleaseMutex(mutex);CloseHandle(mutex);goto failed; }
  }
#ifdef XLL_PROBE_EXTRACT_ONLY
  ReleaseMutex(mutex);CloseHandle(mutex);puts("control_resource_extract_and_sha256_pass");return 0;
#endif
  if(!health()) {
    stage=L"process";
    STARTUPINFOW si={0}; PROCESS_INFORMATION pi={0}; si.cb=sizeof(si);
    WCHAR command[32768];swprintf(command,32768,L"\"%ls\"",exe);
    if(!CreateProcessW(exe,command,0,0,FALSE,CREATE_NO_WINDOW,0,path,&si,&pi)) {ReleaseMutex(mutex);CloseHandle(mutex);goto failed;}
    CloseHandle(pi.hThread);CloseHandle(pi.hProcess);
    for(int i=0;i<60 && !health();i++)Sleep(250);
  }
  ReleaseMutex(mutex);CloseHandle(mutex);
  if(!health())goto failed;
  InterlockedExchange(&running,0);return 0;
failed:
#ifdef XLL_PROBE_EXTRACT_ONLY
  fwprintf(stderr,L"control_failure: %ls / %lu\n",stage,GetLastError());return 1;
#else
  {WCHAR message[256];swprintf(message,256,L"办公桥 XLL 验证失败：%ls / %lu。用户文件未修改。",stage,GetLastError());MessageBoxW(0,message,L"OfficeBridge XLL Probe",MB_OK|MB_ICONERROR);}
  InterlockedExchange(&running,0);return 1;
#endif
}
__declspec(dllexport) int __cdecl xlAutoOpen(void) {
  if(InterlockedCompareExchange(&running,1,0)==0) {
    // Keep this DLL resident while its finite initialization worker is active.
    HMODULE pinned=0;GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS|GET_MODULE_HANDLE_EX_FLAG_PIN,(LPCWSTR)&xlAutoOpen,&pinned);
    if(pinned)module=pinned;
    HANDLE thread=CreateThread(0,0,start,0,0,0);
    if(thread)CloseHandle(thread);else InterlockedExchange(&running,0);
  }
  return 1;
}
__declspec(dllexport) int __cdecl xlAutoAdd(void) { return 1; }
__declspec(dllexport) int __cdecl xlAutoClose(void) { return 1; }
__declspec(dllexport) int __cdecl xlAutoRemove(void) { return 1; }
BOOL WINAPI DllMain(HINSTANCE instance,DWORD reason,LPVOID reserved) {
  (void)reserved;if(reason==DLL_PROCESS_ATTACH){module=instance;DisableThreadLibraryCalls(instance);}return TRUE;
}
