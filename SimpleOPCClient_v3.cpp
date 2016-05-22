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
// KEPWARE´s  OPC client sample code. A few wrapper functions to initiate
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

#include <atlbase.h>    // required for using the "_T" macro
#include <iostream>
#include <ObjIdl.h>

#include "opcda.h"
#include "opcerror.h"
#include "SimpleOPCClient_v3.h"
#include "SOCAdviseSink.h"
#include "SOCDataCallback.h"
#include "SOCWrapperFunctions.h"

using namespace std;

#define OPC_SERVER_NAME L"Matrikon.OPC.Simulation.1"

//#define REMOTE_SERVER_NAME L"your_path"

// Global variables

// The OPC DA Spec requires that some constants be registered in order to use
// them. The one below refers to the OPC DA 1.0 IDataObject interface.
UINT OPC_DATA_TIME = RegisterClipboardFormat (_T("OPCSTMFORMATDATATIME"));

wchar_t *ITEM_READ[3]= {L"Random.Int2",L"Random.Real4",L"Random.Time"};
wchar_t *ITEM_WRITE[3] = {L"Bucket Brigade.Int2", L"Bucket Brigade.Int4", L"Bucket Brigade.String"};
//////////////////////////////////////////////////////////////////////
// Read the value of an item on an OPC server. 
//
void main(void)
{
	IOPCServer* pIOPCServer = NULL;   //pointer to IOPServer interface
	IOPCItemMgt* pIOPCItemRead = NULL; //pointer to IOPCItemMgt interface
	IOPCItemMgt* pIOPCItemWrite = NULL; //pointer to IOPCItemMgt interface for the writing group

	OPCHANDLE hServerRead; // server handle to the reading group
	OPCHANDLE hServerWrite; // server handle to the writing group
	OPCHANDLE hServerReadArray[3];  // server handle to the reading item array
	OPCHANDLE hServerWriteArray[3];  // server handle to the writing item array

	LPCWSTR readingGroupName = L"Group1";
	LPCWSTR writingGroupName = L"Group2";

	int i;
	char buf[100];

	// Have to be done before using microsoft COM library:
	printf("Initializing the COM environment...\n");
	CoInitialize(NULL);

	// Let's instantiante the IOPCServer interface and get a pointer of it:
	printf("Intantiating the MATRIKON OPC Server for Simulation...\n");
	pIOPCServer = InstantiateServer(OPC_SERVER_NAME);
	
	// Add the OPC groups for reading and writing the OPC server and get an handle to the IOPCItemMgt
	//interfaces:
	printf("Adding a reading group in the INACTIVE state for the moment...\n");
	AddTheGroup(pIOPCServer, pIOPCItemRead, hServerRead, readingGroupName);

	printf("Adding a writing group in the INACTIVE state...\n");
	AddTheGroup(pIOPCServer, pIOPCItemWrite, hServerWrite, writingGroupName);


	// Add the OPC items. First we have to convert from wchar_t* to char*
	// in order to print the item name in the console.
	size_t m;
	for (i=0; i<3; i++) {
		wcstombs_s(&m, buf, 100, ITEM_READ[i], _TRUNCATE);
		printf("Adding the item %s to the reading group...\n", buf);
	}

    AddReadingItems(pIOPCItemRead, hServerReadArray);

	for (i = 0; i<3; i++) {
		wcstombs_s(&m, buf, 100, ITEM_WRITE[i], _TRUNCATE);
		printf("Adding the item %s to the writing group...\n", buf);
	}
	AddWritingItems(pIOPCItemWrite, hServerWriteArray);

	int bRet;
	MSG msg;
	DWORD ticks1, ticks2;
    
	// Establish a callback asynchronous read by means of the IOPCDaraCallback
	// (OPC DA 2.0) method. We first instantiate a new SOCDataCallback object and
	// adjusts its reference count, and then call a wrapper function to
	// setup the callback.
	IConnectionPoint* pIConnectionPoint = NULL; //pointer to IConnectionPoint Interface
	DWORD dwCookie = 0;
	SOCDataCallback* pSOCDataCallback = new SOCDataCallback();
	pSOCDataCallback->AddRef();

	printf("Setting up the IConnectionPoint callback connection...\n");
	SetDataCallback(pIOPCItemRead, pSOCDataCallback, pIConnectionPoint, &dwCookie);

	// Change the reading group to the ACTIVE state so that we can receive the
	// server´s callback notification
	printf("Changing the reading group state to ACTIVE...\n");
    SetGroupActive(pIOPCItemRead);

	// Enter a message pump in order to process the server´s callback
	// notifications
		
	ticks1 = GetTickCount();
	printf("Waiting for IOPCDataCallback notifications during 10 seconds...\n");
	do {
		bRet = GetMessage( &msg, NULL, 0, 0 );
		if (!bRet){
			printf ("Failed to get windows message! Error code = %d\n", GetLastError());
			exit(0);
		}
		TranslateMessage(&msg); // This call is not really needed ...
		DispatchMessage(&msg);  // ... but this one is!
        ticks2 = GetTickCount();
	} while ((ticks2 - ticks1) < 10000);

	// Cancel the callback and release its reference
	printf("Cancelling the IOPCDataCallback notifications...\n");
    CancelDataCallback(pIConnectionPoint, dwCookie);
	//pIConnectionPoint->Release();
	pSOCDataCallback->Release();

	// Remove the OPC items:
	printf("Removing the OPC reading items...\n");
	RemoveItem(pIOPCItemRead, hServerReadArray);

	printf("Removing the OPC writing items...\n");
	RemoveItem(pIOPCItemWrite, hServerWriteArray);

	// Remove the OPC group:
	printf("Removing the OPC group objects...\n");
    pIOPCItemRead->Release();
	pIOPCItemWrite->Release();
	RemoveGroup(pIOPCServer, hServerRead);
	RemoveGroup(pIOPCServer, hServerWrite);

	// release the interface references:
	printf("Removing the OPC server object...\n");
	pIOPCServer->Release();

	//close the COM library:
	printf ("Releasing the COM environment...\n");
	CoUninitialize();
}

////////////////////////////////////////////////////////////////////
// Instantiate the IOPCServer interface of the OPCServer
// having the name ServerName. Return a pointer to this interface
//
IOPCServer* InstantiateServer(wchar_t ServerName[])
{
	CLSID CLSID_OPCServer;
	HRESULT hr;

	// get the CLSID from the OPC Server Name:
	hr = CLSIDFromString(ServerName, &CLSID_OPCServer);
	_ASSERT(!FAILED(hr));


	//queue of the class instances to create
	LONG cmq = 1; // nbr of class instance to create.
	MULTI_QI queue[1] =
		{{&IID_IOPCServer,
		NULL,
		0}};

	//Server info:
	//COSERVERINFO CoServerInfo =
    //{
	//	/*dwReserved1*/ 0,
	//	/*pwszName*/ REMOTE_SERVER_NAME,
	//	/*COAUTHINFO*/  NULL,
	//	/*dwReserved2*/ 0
    //}; 

	// create an instance of the IOPCServer
	hr = CoCreateInstanceEx(CLSID_OPCServer, NULL, CLSCTX_SERVER,
		/*&CoServerInfo*/NULL, cmq, queue);
	_ASSERT(!hr);

	// return a pointer to the IOPCServer interface:
	return(IOPCServer*) queue[0].pItf;
}


/////////////////////////////////////////////////////////////////////
// Add group "Group1" to the Server whose IOPCServer interface
// is pointed by pIOPCServer. 
// Returns a pointer to the IOPCItemMgt interface of the added group
// and a server opc handle to the added group.
//
void AddTheGroup(IOPCServer* pIOPCServer, IOPCItemMgt* &pIOPCItemMgt, 
				 OPCHANDLE& hServerGroup, LPCWSTR groupName)
{
	DWORD dwUpdateRate = 0;
	OPCHANDLE hClientGroup = 0;

	// Add an OPC group and get a pointer to the IUnknown I/F:
    HRESULT hr = pIOPCServer->AddGroup(/*szName*/ groupName,
		/*bActive*/ FALSE,
		/*dwRequestedUpdateRate*/ 500,
		/*hClientGroup*/ hClientGroup,
		/*pTimeBias*/ 0,
		/*pPercentDeadband*/ 0,
		/*dwLCID*/0,
		/*phServerGroup*/&hServerGroup,
		&dwUpdateRate,
		/*riid*/ IID_IOPCItemMgt,
		/*ppUnk*/ (IUnknown**) &pIOPCItemMgt);
	_ASSERT(!FAILED(hr));
}



//////////////////////////////////////////////////////////////////
// Add the Item ITEM_ID to the group whose IOPCItemMgt interface
// is pointed by pIOPCItemMgt pointer. Return a server opc handle
// to the item.
 
void AddWritingItems(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE *hServerItem)
{
	HRESULT hr;
	int i;
	// Array of items to add:
	OPCITEMDEF ItemArray[3] = {
		{L"", ITEM_WRITE[0], TRUE, 1, 0, NULL, VT_I2, 0},
		{L"", ITEM_WRITE[1], TRUE, 1, 0, NULL, VT_I4, 0},
		{L"", ITEM_WRITE[2], TRUE, 1, 0, NULL, VT_BSTR, 0}
	};


	//Add Result:
	OPCITEMRESULT* pAddResult=NULL;
	HRESULT* pErrors = NULL;

	// Add an Item to the previous Group:
	hr = pIOPCItemMgt->AddItems(3, ItemArray, &pAddResult, &pErrors);
	if (hr != S_OK){
		printf("Failed call to AddItems function. Error code = %x\n", hr);
		exit(0);
	}

	// Server handle for the added item:
	for (i=0; i<3; i++) {
		hServerItem[i] = pAddResult[i].hServer;
	}


	// release memory allocated by the server:
	CoTaskMemFree(pAddResult->pBlob);

	CoTaskMemFree(pAddResult);
	pAddResult = NULL;

	CoTaskMemFree(pErrors);
	pErrors = NULL;
}

void AddReadingItems(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE *hServerItem)
{
	HRESULT hr;
	int i;
	// Array of items to add:
	OPCITEMDEF ItemArray[3] = {
		{ L"", ITEM_READ[0], TRUE, 1, 0, NULL, VT_I2, 0 },
		{ L"", ITEM_READ[1], TRUE, 1, 0, NULL, VT_R4, 0 },
		{ L"", ITEM_READ[2], TRUE, 1, 0, NULL, VT_DATE, 0 }
	};


	//Add Result:
	OPCITEMRESULT* pAddResult = NULL;
	HRESULT* pErrors = NULL;

	// Add an Item to the previous Group:
	hr = pIOPCItemMgt->AddItems(3, ItemArray, &pAddResult, &pErrors);
	if (hr != S_OK) {
		printf("Failed call to AddItems function. Error code = %x\n", hr);
		exit(0);
	}

	// Server handle for the added item:
	for (i = 0; i<3; i++) {
		hServerItem[i] = pAddResult[i].hServer;
	}


	// release memory allocated by the server:
	CoTaskMemFree(pAddResult->pBlob);

	CoTaskMemFree(pAddResult);
	pAddResult = NULL;

	CoTaskMemFree(pErrors);
	pErrors = NULL;
}

///////////////////////////////////////////////////////////////////////////
// Remove the item whose server handle is hServerItem from the group
// whose IOPCItemMgt interface is pointed by pIOPCItemMgt
//
void RemoveItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE *hServerArray)
{
	
	//Remove the item:
	HRESULT* pErrors; // to store error code(s)
	HRESULT hr = pIOPCItemMgt->RemoveItems(3, hServerArray, &pErrors);
	_ASSERT(!hr);

	//release memory allocated by the server:
	CoTaskMemFree(pErrors);
	pErrors = NULL;
}

////////////////////////////////////////////////////////////////////////
// Remove the Group whose server handle is hServerGroup from the server
// whose IOPCServer interface is pointed by pIOPCServer
//
void RemoveGroup (IOPCServer* pIOPCServer, OPCHANDLE hServerGroup)
{
	// Remove the group:
	HRESULT hr = pIOPCServer->RemoveGroup(hServerGroup, FALSE);
	if (hr != S_OK){
		if (hr == OPC_S_INUSE)
			printf ("Failed to remove OPC group: object still has references to it.\n");
		else printf ("Failed to remove OPC group. Error code = %x\n", hr);
		exit(0);
	}
}

////////////////////////////////////////////////////////////////////////
// Writes the value pointed by varValue to the item whose handle is 
// hServerItem, using the IOPCSyncIO interface obtained
// from the group interface pointed by pGroupIUnknown
//
void WriteItem(IUnknown* pGroupIUnknown, OPCHANDLE hServerItem, VARIANT *varValue)
{
	// value of the item:
	OPCITEMSTATE* pValue = NULL;

	//get a pointer to the IOPCSyncIOInterface:
	IOPCSyncIO* pIOPCSyncIO;
	pGroupIUnknown->QueryInterface(__uuidof(pIOPCSyncIO), (void**) &pIOPCSyncIO);

	// write the item value to the device:
	HRESULT* pErrors = NULL; //to store error code(s)
	HRESULT hr = pIOPCSyncIO->Write(1, &hServerItem, varValue, &pErrors);
	_ASSERT(!hr);
	_ASSERT(pValue!=NULL);


	//Release memeory allocated by the OPC server:
	CoTaskMemFree(pErrors);
	pErrors = NULL;

	CoTaskMemFree(pValue);
	pValue = NULL;

	// release the reference to the IOPCSyncIO interface:
	pIOPCSyncIO->Release();
}
