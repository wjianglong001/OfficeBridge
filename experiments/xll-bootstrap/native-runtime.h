#ifndef OFFICEBRIDGE_NATIVE_RUNTIME_H
#define OFFICEBRIDGE_NATIVE_RUNTIME_H

#include <windows.h>
#include <wincrypt.h>
#include <wchar.h>
#include <string.h>

/* Included by native-server.c after image, stop_server and request_deadline.
 * Only the compiled sibling service is eligible; no caller-supplied path or
 * arguments participate. This code never starts a process. */
static HANDLE native_runtime_file=INVALID_HANDLE_VALUE;
static WCHAR native_runtime_path[32768];
static HANDLE *native_runtime_directories;
static size_t native_runtime_directory_count;

static void native_runtime_release(void) {
  DWORD saved=GetLastError();
  if(native_runtime_file!=INVALID_HANDLE_VALUE) {
    CloseHandle(native_runtime_file);native_runtime_file=INVALID_HANDLE_VALUE;
  }
  while(native_runtime_directory_count) {
    HANDLE directory=native_runtime_directories[--native_runtime_directory_count];
    if(directory!=INVALID_HANDLE_VALUE)CloseHandle(directory);
  }
  if(native_runtime_directories)HeapFree(GetProcessHeap(),0,native_runtime_directories);
  native_runtime_directories=0;native_runtime_path[0]=0;
  SetLastError(saved);
}
static DWORD native_runtime_interrupted(void) {
  if(stop_server)return ERROR_CANCELLED;
  if(GetTickCount64()>=request_deadline)return ERROR_TIMEOUT;
  return 0;
}
static DWORD native_runtime_local_drive(const WCHAR *path) {
  WCHAR drive[3]={path[0],L':',0},root[4]={path[0],L':',L'\\',0},device[1024];
  const WCHAR *prefix=L"\\Device\\HarddiskVolume",*suffix;
  UINT type;
  if(!((path[0]>=L'A'&&path[0]<=L'Z')||(path[0]>=L'a'&&path[0]<=L'z'))||path[1]!=L':'||path[2]!=L'\\')return ERROR_BAD_PATHNAME;
  type=GetDriveTypeW(root);
  if(type!=DRIVE_FIXED&&type!=DRIVE_REMOVABLE)return ERROR_BAD_NETPATH;
  if(!QueryDosDeviceW(drive,device,(DWORD)(sizeof(device)/sizeof(device[0]))))return GetLastError();
  if(wcsncmp(device,prefix,wcslen(prefix)))return ERROR_BAD_NETPATH;
  suffix=device+wcslen(prefix);
  if(*suffix<L'0'||*suffix>L'9')return ERROR_BAD_NETPATH;
  while(*suffix>=L'0'&&*suffix<=L'9')suffix++;
  return *suffix?ERROR_BAD_NETPATH:0;
}
static DWORD native_runtime_plain_path(const WCHAR *path) {
  size_t start=3,index;
  DWORD error=native_runtime_local_drive(path);if(error)return error;
  for(index=3;;index++) {
    WCHAR value=path[index];
    if(!value||value==L'\\') {
      size_t length=index-start;
      if(!length||(length==1&&path[start]==L'.')||(length==2&&path[start]==L'.'&&path[start+1]==L'.')||path[index-1]==L'.'||path[index-1]==L' ')return ERROR_BAD_PATHNAME;
      if(!value)break;
      start=index+1;
    } else if(value<0x20||value==L'/'||value==L':'||value==L'"'||value==L'<'||value==L'>'||value==L'|'||value==L'?'||value==L'*')return ERROR_BAD_PATHNAME;
  }
  return 0;
}
static DWORD native_runtime_same_final_path(HANDLE file,const WCHAR *expected) {
  WCHAR final_path[32768];
  DWORD length=GetFinalPathNameByHandleW(file,final_path,32768,FILE_NAME_NORMALIZED|VOLUME_NAME_DOS);
  if(!length)return GetLastError();
  if(length>=32768)return ERROR_FILENAME_EXCED_RANGE;
  /* Reject UNC, DOS-device aliases, 8.3 spellings, dot components and other
   * normalization differences rather than silently changing launch paths. */
  if(wcsncmp(final_path,L"\\\\?\\",4)||_wcsicmp(final_path+4,expected))return ERROR_BAD_PATHNAME;
  return native_runtime_local_drive(final_path+4);
}
static DWORD native_runtime_hold_directory(const WCHAR *path) {
  HANDLE directory;
  BY_HANDLE_FILE_INFORMATION information;
  DWORD error=native_runtime_interrupted();if(error)return error;
  directory=CreateFileW(path,FILE_READ_ATTRIBUTES,FILE_SHARE_READ|FILE_SHARE_WRITE,0,OPEN_EXISTING,FILE_FLAG_OPEN_REPARSE_POINT|FILE_FLAG_BACKUP_SEMANTICS,0);
  if(directory==INVALID_HANDLE_VALUE)return GetLastError();
  if(!GetFileInformationByHandle(directory,&information))error=GetLastError();
  else if(!(information.dwFileAttributes&FILE_ATTRIBUTE_DIRECTORY)||(information.dwFileAttributes&FILE_ATTRIBUTE_REPARSE_POINT))error=ERROR_INVALID_DATA;
  else error=native_runtime_same_final_path(directory,path);
  if(error)CloseHandle(directory);
  else native_runtime_directories[native_runtime_directory_count++]=directory;
  return error;
}
static DWORD verify_native_runtime(void) {
  DWORD error=native_runtime_interrupted();
  DWORD length,count=0,size;
  WCHAR *separator,saved;
  size_t prefix,index;
  BY_HANDLE_FILE_INFORMATION information;
  HCRYPTPROV provider=0;
  HCRYPTHASH hash=0;
  BYTE buffer[65536],digest[32];
  if(error)return error;
  if(native_runtime_file!=INVALID_HANDLE_VALUE)return 0;
  native_runtime_release();
  length=GetModuleFileNameW(image,native_runtime_path,32768);
  if(!length||length>=32768){error=ERROR_BAD_PATHNAME;goto cleanup;}
  separator=wcsrchr(native_runtime_path,L'\\');
  if(!separator){error=ERROR_BAD_PATHNAME;goto cleanup;}
  prefix=(size_t)(separator-native_runtime_path)+1;
  if(prefix+24>32768){error=ERROR_FILENAME_EXCED_RANGE;goto cleanup;}
  wcscpy(native_runtime_path+prefix,L"OfficeBridgeService.exe");
  error=native_runtime_plain_path(native_runtime_path);if(error)goto cleanup;
  /* There can be at most one directory component per two characters.
   * Keep every ancestor, including the volume root, open without DELETE
   * sharing until release so it cannot be renamed or replaced. */
  native_runtime_directories=(HANDLE*)HeapAlloc(GetProcessHeap(),0,(prefix/2+1)*sizeof(HANDLE));
  if(!native_runtime_directories){error=ERROR_NOT_ENOUGH_MEMORY;goto cleanup;}
  saved=native_runtime_path[3];native_runtime_path[3]=0;
  error=native_runtime_hold_directory(native_runtime_path);native_runtime_path[3]=saved;
  if(error)goto cleanup;
  for(index=3;index<prefix;index++) {
    if(native_runtime_path[index]!=L'\\')continue;
    saved=native_runtime_path[index];native_runtime_path[index]=0;
    error=native_runtime_hold_directory(native_runtime_path);native_runtime_path[index]=saved;
    if(error)goto cleanup;
  }
  error=native_runtime_interrupted();if(error)goto cleanup;
  native_runtime_file=CreateFileW(native_runtime_path,GENERIC_READ,FILE_SHARE_READ,0,OPEN_EXISTING,FILE_FLAG_OPEN_REPARSE_POINT|FILE_FLAG_SEQUENTIAL_SCAN,0);
  if(native_runtime_file==INVALID_HANDLE_VALUE){error=GetLastError();goto cleanup;}
  if(!GetFileInformationByHandle(native_runtime_file,&information)){error=GetLastError();goto cleanup;}
  if(information.dwFileAttributes&(FILE_ATTRIBUTE_REPARSE_POINT|FILE_ATTRIBUTE_DIRECTORY)){error=ERROR_INVALID_DATA;goto cleanup;}
  error=native_runtime_same_final_path(native_runtime_file,native_runtime_path);if(error)goto cleanup;
  if(!CryptAcquireContextW(&provider,0,0,PROV_RSA_AES,CRYPT_VERIFYCONTEXT)||!CryptCreateHash(provider,CALG_SHA_256,0,0,&hash)){error=GetLastError();goto cleanup;}
  for(;;) {
    error=native_runtime_interrupted();if(error)goto cleanup;
    if(!ReadFile(native_runtime_file,buffer,sizeof(buffer),&count,0)){error=GetLastError();goto cleanup;}
    if(!count)break;
    if(!CryptHashData(hash,buffer,count,0)){error=GetLastError();goto cleanup;}
  }
  size=sizeof(digest);
  if(!CryptGetHashParam(hash,HP_HASHVAL,digest,&size,0)){error=GetLastError();goto cleanup;}
  if(size!=sizeof(digest)||memcmp(digest,NATIVE_BACKEND_SHA256,sizeof(digest))){error=ERROR_CRC;goto cleanup;}
  error=native_runtime_interrupted();
cleanup:
  if(hash)CryptDestroyHash(hash);
  if(provider)CryptReleaseContext(provider,0);
  if(error)native_runtime_release();
  return error;
}
#endif
