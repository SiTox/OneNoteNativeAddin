/*!-----------------------------------------------------------------------
	ids.h

	Contains all the progids/guids/clsids/strings used across the addin. If an
	id is needed to be updated then it should only need to be changed
	here and nowhere else.

	Remember these ids need to be unique on each system so they should be
	changed if this sample addin is used as a base for a another addin.

-----------------------------------------------------------------------!*/
#pragma once

// Strings
#define ADDIN_PROGID            OLESTR("KeepOneNoteRunningCOMAddin.Connect")

// GUIDs
#define TYPELIB_GUID            511AEFF8-CE96-46DC-B138-A024D8580BBF
#define ADDIN_CLSID             81141740-CBE1-42A4-AD7D-4F1C6AF4F06D

// Remember to keep these string version in sync with the ones above
#define TYPELIB_GUID_STR        OLESTR("511AEFF8-CE96-46DC-B138-A024D8580BBF")
#define ADDIN_CLSID_STR         OLESTR("81141740-CBE1-42A4-AD7D-4F1C6AF4F06D")

