// Simple OPC Client
//
// This is a modified version of the "Simple OPC Client" originally
// developed by Philippe Gras (CERN) for demonstrating the basic techniques
// involved in the development of an OPC DA client.
//
// The modifications are the introduction of two C++ classes to allow the
// the client to ask for callback notifications from the OPC server, and
// the corresponding introduction of a message comsumption loop in the
// main program to allow the client to process those notifications. The
// C++ classes implement the OPC DA 1.0 IAdviseSink and the OPC DA 2.0
// IOPCDataCallback client interfaces, and in turn were adapted from the
// KEPWARE�s  OPC client sample code. A few wrapper functions to initiate
// and to cancel the notifications were also developed.
//
// The original Simple OPC Client code can still be found (as of this date)
// in
//        http://pgras.home.cern.ch/pgras/OPCClientTutorial/
//
//
// Luiz T. S. Mendes - DELT/UFMG - 15 Sept 2011
// luizt at cpdee.ufmg.br
//

#ifndef SIMPLE_OPC_CLIENT_H
#define SIMPLE_OPC_CLIENT_H


struct readStruct {
	int prod;
	float oee;
	char time[8];
};

struct writeStruct {
	int cim;
	int ton;
	char time[8];
};

IOPCServer *InstantiateServer(wchar_t ServerName[]);
void AddTheGroup(IOPCServer* pIOPCServer, IOPCItemMgt* &pIOPCItemMgt,
	OPCHANDLE& hServerGroup, LPCWSTR groupName);
void AddWritingItems(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE *hServerItem);
void AddReadingItems(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE *hServerItem);
void RemoveItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE *hServerArray);
void RemoveGroup(IOPCServer* pIOPCServer, OPCHANDLE hServerGroup);
void WriteItem(IUnknown * pGroupIUnknown, OPCHANDLE hServerItem, VARIANT * varValue);
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,
	LPSTR lpCmdLine, int nCmdShow);
DWORD WINAPI OPCThread1(LPVOID id);
DWORD WINAPI SocketThread(LPVOID id);
DWORD WINAPI GuiThread(LPVOID hInst);
void DataChanged(VARIANT pValue, char* value);
#endif // SIMPLE_OPC_CLIENT_H not defined
