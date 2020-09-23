// User Tab 2 Hook DLL:
// (c)1999, 2000, 2001, 2004 by Louis. Unauthorized copying prohibited.
//
// Downloaded from www.louis-coder.com.
// Use this dll in combination with GFGlobalKeyHook to set up
// a global key hook.
//
// ***INCLUDES***

#include "StdAfx.h"
#include "kh.h"
#include "winuser.h"

// ***VARIABLES***

#pragma data_seg("SHARDATA")
// Key Hook
int KeyHookTemp = 0;
// Key Hook constants
int KEYHOOK_SHIFT = 1;
int KEYHOOK_STRG = 2;
int KEYHOOK_ALT = 4;
// other
HHOOK KeyHookHandle = 0;
HWND MsgTargetAddress = 0; // target address of message
#pragma data_seg()
HINSTANCE Inst;

// ***PROTOTYPES***

void    _stdcall noname_sub001(HWND MsgTargetAddress);
void    _stdcall noname_sub002();
LRESULT _stdcall noname_sub003(int nCode, WPARAM wParam, LPARAM lParam);

// ***CODE***

BOOL WINAPI DllMain (HANDLE hInst, ULONG ul_reason_for_call, LPVOID lpReserved)
{
	Inst = hInst;
	return 1;
}

void _stdcall noname_sub001(HWND MsgTargetAddressPassed, LPCTSTR HookDLLName)
{
	MsgTargetAddress = MsgTargetAddressPassed; // MsgTargetAddress is handle of receiving object
	KeyHookHandle = SetWindowsHookEx(WH_KEYBOARD, (HOOKPROC)noname_sub003, GetModuleHandle(HookDLLName), 0); // pass HookDLLName to allow renaming this dll
}

void _stdcall noname_sub002()
{
	UnhookWindowsHookEx(KeyHookHandle);
}

LRESULT _stdcall noname_sub003(int nCode, WPARAM wParam, LPARAM lParam)
{
	if (nCode == HC_ACTION) // check if key was repeated in VB
	{
		KeyHookTemp = 0; // reset
		if ((GetKeyState(20) & 0x0001) != 0) // check if CapsLock is pressed; use GetKeyState() to get key status
		{
			KeyHookTemp = KeyHookTemp + KEYHOOK_SHIFT;
		}
		else
		{
			if (GetAsyncKeyState(16) != 0) // use GetAsyncKeyState() to get current key state
				KeyHookTemp = KeyHookTemp + KEYHOOK_SHIFT;
		}
		if (GetAsyncKeyState(17) != 0) 
			KeyHookTemp = KeyHookTemp + KEYHOOK_STRG;
		if (GetAsyncKeyState(18) != 0)
			KeyHookTemp = KeyHookTemp + KEYHOOK_ALT;
		
		SendMessage(MsgTargetAddress, 0x0000, wParam, KeyHookTemp);
	
	}
	return CallNextHookEx(KeyHookHandle, nCode, wParam, lParam);
}
