#ifndef OFFICEBRIDGE_NATIVE_REGISTRY_H
#define OFFICEBRIDGE_NATIVE_REGISTRY_H

#include <stddef.h>
#include <string.h>
#include "vendor/cJSON/cJSON.h"

#if CJSON_NESTING_LIMIT != 32
#error Compile both the consumer and cJSON.c with -DCJSON_NESTING_LIMIT=32
#endif

/* cJSON stores strings as C strings: reject decoded NUL before parsing.
 * Escaped backslashes remain ordinary text; "\\u0000" is not a NUL.
 * The caller provides a complete NUL-terminated response, rejecting actual
 * embedded NUL bytes at the transport boundary. Never log this response. */
static int officebridge_registry_strings_safe(const char *json) {
  const unsigned char *p=(const unsigned char *)json;
  int quoted=0;
  while(*p) {
    unsigned char value=*p++;
    if(!quoted) {
      if(value=='"')quoted=1;
      else if(value<0x20&&value!='\t'&&value!='\r'&&value!='\n')return 0;
      continue;
    }
    if(value=='"'){quoted=0;continue;}
    if(value<0x20)return 0;
    if(value!='\\')continue;
    value=*p++;if(!value)return 0;
    if(value=='u') {
      unsigned int i;
      for(i=0;i<4;i++) {
        unsigned char digit=p[i];
        if(!digit||!((digit>='0'&&digit<='9')||(digit>='a'&&digit<='f')||(digit>='A'&&digit<='F')))return 0;
      }
      if(p[0]=='0'&&p[1]=='0'&&p[2]=='0'&&p[3]=='0')return 0;
      p+=4;
    } else if(value!='"'&&value!='\\'&&value!='/'&&value!='b'&&value!='f'&&value!='n'&&value!='r'&&value!='t')return 0;
  }
  return !quoted;
}

/* Compare decoded names so escaped spellings cannot bypass duplicate checks.
 * Validate the whole parsed tree, including unrelated add-in metadata. */
static int officebridge_registry_tree_safe(const cJSON *node,unsigned int depth) {
  const cJSON *child,*other;
  if(!node||depth>32)return 0;
  if(cJSON_IsObject(node)) {
    for(child=node->child;child;child=child->next) {
      if(!child->string)return 0;
      for(other=node->child;other!=child;other=other->next)
        if(!strcmp(child->string,other->string))return 0;
    }
  }
  for(child=node->child;child;child=child->next)
    if(!officebridge_registry_tree_safe(child,depth+1))return 0;
  return 1;
}
static int officebridge_registry_exact(const cJSON *record,const char *key,const char *expected) {
  const cJSON *field=cJSON_GetObjectItemCaseSensitive(record,key);
  return cJSON_IsString(field)&&field->valuestring&&!strcmp(field->valuestring,expected);
}

/* 0: no matching name; 1: exactly one trusted, matching enabled registration;
 * -1: unverified, malformed, conflicting, disabled or ambiguous. Read-only. */
static int officebridge_registry_match_internal(const char *json,const char *expected_name,const char *expected_url,const char *expected_version,int allow_empty_online_version) {
  cJSON *root;
  const cJSON *record;
  const char *end=0;
  size_t length=0;
  int found=0,result=-1;
  if(!json||!expected_name||!*expected_name||!expected_url||!*expected_url||!expected_version||!*expected_version)return -1;
  while(length<=65536&&json[length])length++;
  if(!length||length>65536||!officebridge_registry_strings_safe(json))return -1;
  root=cJSON_ParseWithLengthOpts(json,length+1,&end,1);
  if(!root)return -1;
  if(!cJSON_IsArray(root)||!end||*end||!officebridge_registry_tree_safe(root,0))goto cleanup;
  for(record=root->child;record;record=record->next) {
    const cJSON *name;
    if(!cJSON_IsObject(record))goto cleanup;
    name=cJSON_GetObjectItemCaseSensitive(record,"name");
    if(!cJSON_IsString(name)||!name->valuestring)goto cleanup;
    if(strcmp(name->valuestring,expected_name))continue;
    if(found++)goto cleanup;
    if(!officebridge_registry_exact(record,"addonType","et")||
       !officebridge_registry_exact(record,"online","true")||
       !officebridge_registry_exact(record,"enable","enable_dev")||
       !officebridge_registry_exact(record,"url",expected_url)||
       (!officebridge_registry_exact(record,"version",expected_version)&&
        !(allow_empty_online_version&&officebridge_registry_exact(record,"version",""))))goto cleanup;
  }
  result=found?1:0;
cleanup:
  cJSON_Delete(root);
  return result;
}
/* Strict callers, including the probe, retain exact version semantics. */
static int officebridge_registry_match(const char *json,const char *expected_name,const char *expected_url,const char *expected_version) {
  return officebridge_registry_match_internal(json,expected_name,expected_url,expected_version,0);
}
/* Explicit FULL-bootstrap compatibility: WPS may report an empty string for
 * online versions. All other identity/state fields remain exact; absent,
 * non-string, duplicate or wrong nonempty versions are never accepted. */
static int officebridge_registry_match_online(const char *json,const char *expected_name,const char *expected_url,const char *expected_version) {
  return officebridge_registry_match_internal(json,expected_name,expected_url,expected_version,1);
}
/* Retain the isolated probe namespace and its established contract. */
static int officebridge_registry_state(const char *json) {
  return officebridge_registry_match(json,"OfficeBridgeXllProbe","http://127.0.0.1:43191/probe/","5.0.0");
}
#endif
