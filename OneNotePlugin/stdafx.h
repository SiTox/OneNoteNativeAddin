// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently,
// but are changed infrequently

#pragma once

#ifndef STRICT
#define STRICT
#endif

#include "targetver.h"

#define _ATL_APARTMENT_THREADED
#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// some CString constructors will be explicit

#include <afxwin.h>
#ifndef _AFX_NO_OLE_SUPPORT
#include <afxdisp.h>        // MFC Automation classes
#endif // _AFX_NO_OLE_SUPPORT

#include "resource.h"
#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>

using namespace ATL;

#include <algorithm>
#include <list>

using namespace std;
using namespace std::tr1;

// Import for IDTExtensibility2

//#import "libid:00020430-0000-0000-C000-000000000046"\
//	auto_rename auto_search raw_interfaces_only rename_namespace("stdole")\
//	rename("IPictureDisp", "__IPictureDisp")\
//	rename("DISPPARAMS", "__DISPPARAMS")\
//	rename("EXCEPINFO", "__EXCEPINFO")
//using namespace stdole;

#import "libid:AC0714F2-3D04-11D1-AE7D-00A0C90F26F4"\
	auto_rename auto_search raw_interfaces_only rename_namespace("AddinDesign")
using namespace AddinDesign;

// Office type library (i.e. mso.dll)
#import "libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52"\
	auto_rename auto_search raw_interfaces_only rename_namespace("Office")\
	rename("DocumentProperties", "__DocumentProperties")\
	rename("SearchPath", "__SearchPath")\
	rename("RGB", "__RGB")
using namespace Office;

// OneNote type library (i.e. msoutl.olb)o
//#import "libid:0EA692EE-BB50-4E3C-AEF0-356D91732725"\
//	auto_rename auto_search raw_interfaces_only rename_namespace("OneNote")\
//using namespace OneNote;

// This addin type library
#import "x64\Debug\KeepOneNoteRunning.tlb"\
	auto_rename auto_search raw_interfaces_only rename_namespace("KeepOneNoteRunningAddin")
using namespace KeepOneNoteRunningAddin;

//extern int g_verOLMajor;
