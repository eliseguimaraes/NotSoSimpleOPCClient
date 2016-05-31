// Simple OPC Client
//
// This is a modified version of the "Simple OPC Client" originally
// developed by Philippe Gras (CERN) for demonstrating the basic techniques
// involved in the development of an OPC DA client, and furtherly modified by UFMG's professor Luiz T. S. Mendes.
//
// The modifications made by prof. Luiz are the introduction of two C++ classes to allow the
// the client to ask for callback notifications from the OPC server, and
// the corresponding introduction of a message comsumption loop in the
// main program to allow the client to process those notifications. The
// C++ classes implement the OPC DA 1.0 IAdviseSink and the OPC DA 2.0
// IOPCDataCallback client interfaces, and in turn were adapted from the
// KEPWARE´s  OPC client sample code. A few wrapper functions to initiate
// and to cancel the notifications were also developed.
//
// Further to prof. Luiz's modifications, my modifications were:
// -Removal of the C++ classes that implemented the OPC DA 1.0 IAdviseSink
// -Removal of the C++ class that implemented synchronous reading
// -Introduction of 6 new items, 3 of them to be read and 3 to be written on
// -Introduction of a C++ class that implements synchronous writing using the IOPCSyncIO interface
// -Initialization of the environment using multi thread apartments, instead of the original single thread apartments
// -Removal of the message comsumption loop, no longer necessary with the MTA newly introduced approach
// -Replacement of the C++ class AddTheItems by two similar classes, one for the reading itens and another for the writing ones, both capable of adding arrays of items instead of single items.
// -Changes to the AddTheGroup C++ class, for it to take the group name as a parameter instead of a fixed one and changing the reading interval to 500ms.
// -Creation of two groups - one for reading and another for writing -, both with multiple items.
// -Addition of case VT_DATE to the method VarToStr, with proper conversion between these types implemented as well.
// -Introduction of multithread behavior to the code, transfering all the OPC calls to the OPCThread method.
// -Introduction of mutexes and semaphores to deal with multiple access to variables
// -Integration of a socket server to the code, inside it's own thread, which runs the method SocketThread
//
// The original Simple OPC Client code can still be found (as of this date)
// in
//        http://pgras.home.cern.ch/pgras/OPCClientTutorial/
// 
// 
//
// Elise Guimaraes de Araujo - UFMG - 26/05/2016
// elise.guimaraes at ufmg.br
//

#undef UNICODE

#define WIN32_LEAN_AND_MEAN

#include <windows.h>
#include <winsock2.h>
#include <ws2tcpip.h>
#include <stdlib.h>
#include <stdio.h>
#include <atlbase.h>    // required for using the "_T" macro
#include <iostream>
#include <ObjIdl.h>
#include <comutil.h>	//ConvertStringToBSTR
#include "opcda.h"
#include "opcerror.h"
#include "TP1.h"
#include "SOCAdviseSink.h"
#include "SOCDataCallback.h"
#include "SOCWrapperFunctions.h"

#pragma comment(lib, "comsuppw.lib")
#pragma comment(lib, "kernel32.lib")

// Need to link with Ws2_32.lib
#pragma comment (lib, "Ws2_32.lib")
// #pragma comment (lib, "Mswsock.lib")

#define DEFAULT_BUFLEN 512
#define DEFAULT_PORT "4660"
#define MAX_SEQ 99999999

typedef unsigned (WINAPI *CAST_FUNCTION)(LPVOID);

typedef unsigned *CAST_LPDWORD;

using namespace std;

#define OPC_SERVER_NAME L"Matrikon.OPC.Simulation.1"	// Server name

//#define REMOTE_SERVER_NAME L"your_path"

// Global variables

// The OPC DA Spec requires that some constants be registered in order to use
// them. The one below refers to the OPC DA 1.0 IDataObject interface.
UINT OPC_DATA_TIME = RegisterClipboardFormat(_T("OPCSTMFORMATDATATIME"));

wchar_t *ITEM_READ[3] = { L"Random.Int2",L"Random.Real4",L"Random.Time" };	// Ids of items to be read
wchar_t *ITEM_WRITE[3] = { L"Bucket Brigade.Int2", L"Bucket Brigade.Int4", L"Bucket Brigade.String" };	// Ids of items to be written on

struct writeStruct writeData;
struct readStruct readData;

HANDLE hItemsToWrite;	// Binary Semaphore to indicate new items to be written to the OPC Server

HANDLE thread2;			// Handle 
DWORD dwThreadId;

HANDLE readMutex;


//////////////////////////////////////////////////////////////////////
// Main method. Initializes all the threads and synchronization items.
//
void main(void)
{
	HANDLE hThreads[2]; // Handles for the threads
	DWORD dwThreadId1, dwThreadId2;

	hThreads[0] = (HANDLE)_beginthreadex(NULL, 0, (CAST_FUNCTION)OPCThread1, (LPVOID)0, 0, (CAST_LPDWORD)&dwThreadId1);
	if (hThreads[0]) {
		printf("OPC thread criada com id %0x\n", dwThreadId1);
	}
	else {
		printf("Erro criando thread");
		exit(0);
	}

	hThreads[1] = (HANDLE)_beginthreadex(NULL, 0, (CAST_FUNCTION)SocketThread, (LPVOID)0, 0, (CAST_LPDWORD)&dwThreadId2);
	if (hThreads[1]) {
		printf("OPC thread criada com id %0x\n", dwThreadId1);
	}
	else {
		printf("Erro criando thread");
		exit(0);
	}

	hItemsToWrite = CreateSemaphore(NULL, 0, 1, "WriteSemaphore");

	readMutex = CreateMutex(NULL, FALSE, "ReadMutex");

	WaitForMultipleObjects(2, hThreads, true, INFINITE);

	// Close all kernel objects' handles
	for (int i = 0; i < 2; i++) {
		CloseHandle(hThreads[i]);
	}
	CloseHandle(readMutex);
	CloseHandle(hItemsToWrite);
}

// Method ran by the OPC Thread, responsible for all the OPC-related calls
DWORD WINAPI OPCThread1(LPVOID id) {
	IOPCServer* pIOPCServer = NULL;   //pointer to IOPServer interface

	IOPCItemMgt* pIOPCItemMgtRead = NULL; //pointer to IOPCItemMgt interface
	IOPCItemMgt* pIOPCItemMgtWrite = NULL; //pointer to IOPCItemMgt interface for the writing group

	OPCHANDLE hServerGroupRead; // server handle to the reading group
	OPCHANDLE hServerGroupWrite; // server handle to the writing group

	OPCHANDLE hServerReadArray[3];  // server handle to the reading item array
	OPCHANDLE hServerWriteArray[3];  // server handle to the writing item array

	LPCWSTR readingGroupName = L"Group1";
	LPCWSTR writingGroupName = L"Group2";

	int i;
	char buf[100];

	// Have to be done before using microsoft COM library:
	printf("Inicializando o ambiente COM...\n");
	CoInitializeEx(NULL, COINIT_MULTITHREADED);

	// Let's instantiante the IOPCServer interface and get a pointer of it:
	printf("Instanciando o Matrikon OPC Simulation Server...\n");
	pIOPCServer = InstantiateServer(OPC_SERVER_NAME);

	// Add the OPC groups for reading and writing the OPC server and get an handle to the IOPCItemMgt
	//interfaces:
	printf("Adicionando um grupo de leitura em estado inativo, inicialmente...\n");
	AddTheGroup(pIOPCServer, pIOPCItemMgtRead, hServerGroupRead, readingGroupName);

	printf("Adicionando um grupo de escrita em estado inativo...\n");
	AddTheGroup(pIOPCServer, pIOPCItemMgtWrite, hServerGroupWrite, writingGroupName);


	// Add the OPC items. First we have to convert from wchar_t* to char*
	// in order to print the item name in the console.
	size_t m;
	for (i = 0; i < 3; i++) {
		wcstombs_s(&m, buf, 100, ITEM_READ[i], _TRUNCATE);
		printf("Adicionando o item %s ao grupo de leitura...\n", buf);
	}

	AddReadingItems(pIOPCItemMgtRead, hServerReadArray);

	for (i = 0; i < 3; i++) {
		wcstombs_s(&m, buf, 100, ITEM_WRITE[i], _TRUNCATE);
		printf("Adicionando o item %s ao grupo de escrita...\n", buf);
	}
	AddWritingItems(pIOPCItemMgtWrite, hServerWriteArray);

	// Establish a callback asynchronous read by means of the IOPCDataCallback
	// (OPC DA 2.0) method. We first instantiate a new SOCDataCallback object and
	// adjusts its reference count, and then call a wrapper function to
	// setup the callback.
	IConnectionPoint* pIConnectionPoint = NULL; //pointer to IConnectionPoint Interface
	DWORD dwCookie = 0;
	SOCDataCallback* pSOCDataCallback = new SOCDataCallback();
	pSOCDataCallback->AddRef();

	printf("Configurando a conexão de callback IConnectionPoint...\n");
	SetDataCallback(pIOPCItemMgtRead, pSOCDataCallback, pIConnectionPoint, &dwCookie);

	// Change the reading group to the ACTIVE state so that we can receive the
	// server´s callback notification
	printf("Alterando o estado do grupo de leitura para ATIVO...\n");
	SetGroupActive(pIOPCItemMgtRead);

	printf("Aguardando notificações pela interface IOPCDataCallback...\n");
	VARIANT var;
	DWORD rt;
	while (1) {
		rt = WaitForSingleObject(hItemsToWrite, INFINITE);
		if (rt != WAIT_OBJECT_0) {
			cout << "Ocorreu um erro ao aguardar um semáforo. Encerrando execução...";
			exit(0);
		}
		V_VT(&var) = VT_I2;
		V_I2(&var) = writeData.cim;
		WriteItem(pIOPCItemMgtWrite, hServerWriteArray[0], &var);
		V_VT(&var) = VT_I4;
		V_I4(&var) = writeData.ton;
		WriteItem(pIOPCItemMgtWrite, hServerWriteArray[1], &var);
		V_VT(&var) = VT_BSTR;
		BSTR bstrText = _com_util::ConvertStringToBSTR(writeData.time);
		V_BSTR(&var) = bstrText;
		SysFreeString(bstrText);
		WriteItem(pIOPCItemMgtWrite, hServerWriteArray[2], &var);

	}

	// Cancel the callback and release its reference
	printf("Cancelando as notificações de IOPCDataCallback...\n");
	CancelDataCallback(pIConnectionPoint, dwCookie);
	//pIConnectionPoint->Release();
	pSOCDataCallback->Release();

	// Remove the OPC items:
	printf("Removendo os itens de leitura OPC...\n");
	RemoveItem(pIOPCItemMgtRead, hServerReadArray);

	printf("Removendo os itens de escrita OPC...\n");
	RemoveItem(pIOPCItemMgtWrite, hServerWriteArray);

	// Remove the OPC group:
	printf("Removendo os grupos OPC...\n");
	pIOPCItemMgtRead->Release();
	pIOPCItemMgtWrite->Release();
	RemoveGroup(pIOPCServer, hServerGroupRead);
	RemoveGroup(pIOPCServer, hServerGroupWrite);

	// release the interface references:
	printf("Removendo o objeto servidor...\n");
	pIOPCServer->Release();

	//close the COM library:
	printf("Liberando o ambiente COM...\n");
	CoUninitialize();
	_endthreadex(0);
	return 0;
}


DWORD WINAPI SocketThread(LPVOID id) {
	LONG lOldValue;
	WSADATA wsaData;
	int iResult;

	SOCKET ListenSocket = INVALID_SOCKET;
	SOCKET ClientSocket = INVALID_SOCKET;

	struct addrinfo *result = NULL;
	struct addrinfo hints;

	int iSendResult;
	char recvbuf[DEFAULT_BUFLEN];
	char sendbuf[DEFAULT_BUFLEN];
	int recvbuflen = DEFAULT_BUFLEN;
	int sendbuflen = 0;

	int mCode, mSeq, mType, mTon, ackSeq = 0;
	int m2seq;
	m2seq = 0;
	char mHour[256];
	char ack[256];

	// Initialize Winsock
	iResult = WSAStartup(MAKEWORD(2, 2), &wsaData);
	if (iResult != 0) {
		printf("WSAStartup falhou: %d\n", iResult);
		return 1;
	}
	// configuração da struct addrinfo
	ZeroMemory(&hints, sizeof(hints));
	hints.ai_family = AF_INET;			// Família IPv4
	hints.ai_socktype = SOCK_STREAM;	// Tipo: Stream socket (TCP)
	hints.ai_protocol = IPPROTO_TCP;	// Protocolo: TCP
	hints.ai_flags = AI_PASSIVE;		// Socker passivo (servidor)

										// Endereço e porta do servidor
	iResult = getaddrinfo(NULL, DEFAULT_PORT, &hints, &result);
	if (iResult != 0) {
		printf("getaddrinfo falhou: %d\n", iResult);
		WSACleanup();
		return 1;
	}
	// Cria o socket passivo
	ListenSocket = socket(result->ai_family, result->ai_socktype, result->ai_protocol);
	if (ListenSocket == INVALID_SOCKET) {
		printf("socket falhou: %ld\n", WSAGetLastError());
		freeaddrinfo(result);
		WSACleanup();
		return 1;
	}
	// Linka o socket criado à porta indicada
	iResult = bind(ListenSocket, result->ai_addr, (int)result->ai_addrlen);
	if (iResult == SOCKET_ERROR) {
		printf("bind falhou: %d\n", WSAGetLastError());
		freeaddrinfo(result);
		closesocket(ListenSocket);
		WSACleanup();
		return 1;
	}

	// Libera a estrutura addrinfo, que já não será necessária desse ponto adiante
	freeaddrinfo(result);

	// Fica na escuta por conexões - função bloqueante
	iResult = listen(ListenSocket, SOMAXCONN);
	if (iResult == SOCKET_ERROR) {
		printf("listen falhou: %d\n", WSAGetLastError());
		closesocket(ListenSocket);
		WSACleanup();
		return 1;
	}

	// Aceita o connection socket
	ClientSocket = accept(ListenSocket, NULL, NULL);
	if (ClientSocket == INVALID_SOCKET) {
		printf("accept falhou: %d\n", WSAGetLastError());
		closesocket(ListenSocket);
		WSACleanup();
		return 1;
	}

	// Socket passivo já não é necessário, já que connection socket foi obtido
	closesocket(ListenSocket);

	// Recebe até que o cliente encerre a conexão
	do {

		iResult = recv(ClientSocket, recvbuf, recvbuflen, 0);
		if (iResult > 0) {
			recvbuf[iResult] = '\0'; // Removendo lixo de memória
			printf("\nMensagem recebida: %s \n", recvbuf);
			sscanf(recvbuf, "%d", &mCode);
			if (mCode == 1 && iResult == 44) { //setup
				sscanf(recvbuf + 9, "%d", &mSeq);
				sscanf(recvbuf + 2 * 9, "%d", &mType);
				sscanf(recvbuf + 3 * 9, "%d", &mTon);
				sscanf(recvbuf + 4 * 9, "%s", &mHour);
				mHour[8] = '\0';
				writeData.cim = mType;
				strcpy(writeData.time, mHour);
				writeData.ton = mTon;
				ReleaseSemaphore(hItemsToWrite, 1, &lOldValue);
				//send ack
				sprintf(ack, "%.08d|%.08d", 2, ackSeq);
				ackSeq = (ackSeq + 1)%MAX_SEQ;
				ack[17] = '\0';
				iSendResult = send(ClientSocket, ack, strlen(ack), 0);
				printf("\n Mensagem enviada: %s \n", ack);
				if (iSendResult == SOCKET_ERROR) {
					printf("send falhou: %d\n", WSAGetLastError());
					closesocket(ClientSocket);
					WSACleanup();
					return 1;
				}
			}
			else if (mCode == 5 && iResult == 17) { //solicitação
				WaitForSingleObject(readMutex, INFINITE);

				sprintf(sendbuf, "%.08d|%.08d|%.08d|%08.7g|%s", 10, m2seq, readData.prod, readData.oee, readData.time);
				m2seq = (m2seq + 1) % MAX_SEQ;
				sendbuf[44] = '\0';
				iSendResult = send(ClientSocket, sendbuf, strlen(sendbuf), 0);
				printf("\n Mensagem enviada: %s \n", sendbuf);
				if (iSendResult == SOCKET_ERROR) {
					printf("send falhou: %d\n", WSAGetLastError());
					closesocket(ClientSocket);
					WSACleanup();
					return 1;
				}
				ReleaseMutex(readMutex);

			}
			else {
				printf("A mensagem recebida não é compatível. Encerrando...");
				closesocket(ClientSocket);
				WSACleanup();
				return 1;
			}
		}
		else if (iResult == 0)
			printf("Conexão encerrando...\n");
		else {
			printf("recv falhou: %d\n", WSAGetLastError());
			closesocket(ClientSocket);
			WSACleanup();
			return 1;
		}

	} while (iResult > 0);

	// encerra a conexão
	iResult = shutdown(ClientSocket, SD_SEND);
	if (iResult == SOCKET_ERROR) {
		printf("shutdown falhou com erro: %d\n", WSAGetLastError());
		closesocket(ClientSocket);
		WSACleanup();
		return 1;
	}

	// Limpeza
	closesocket(ClientSocket);
	WSACleanup();
	_endthreadex(0);
	return 0;
}

////////////////////////////////////////////////////////////////////
// Method called by the VarToStr method, indicating new data was
// read assynchronously from the OPC Server.
//
void DataChanged(VARIANT pvar, char* value) {
	WaitForSingleObject(readMutex, INFINITE);
	switch (pvar.vt) {
	case VT_I2:
		readData.prod = pvar.intVal;
		break;
	case VT_R4:
		readData.oee = pvar.fltVal;
		break;
	case VT_DATE:
		SYSTEMTIME s;
		VariantTimeToSystemTime(pvar.date, &s);
		sprintf(readData.time,
			"%02d:%02d:%02d",
			s.wHour,
			s.wMinute,
			s.wSecond);
		break;
	default:
		cout << "Ocorreu um erro ao receber dados assincronamente do OPC Server. Abortando...";
		exit(0);
	}
	ReleaseMutex(readMutex);
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
	{ {&IID_IOPCServer,
	NULL,
	0} };

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
	return(IOPCServer*)queue[0].pItf;
}


/////////////////////////////////////////////////////////////////////
// Add group with name groupName to the Server whose IOPCServer interface
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
		/*ppUnk*/ (IUnknown**)&pIOPCItemMgt);
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
	OPCITEMRESULT* pAddResult = NULL;
	HRESULT* pErrors = NULL;

	// Add an Item to the previous Group:
	hr = pIOPCItemMgt->AddItems(3, ItemArray, &pAddResult, &pErrors);
	if (hr != S_OK) {
		printf("Chama falha a AddItems. Error code = %x\n", hr);
		exit(0);
	}

	// Server handle for the added item:
	for (i = 0; i < 3; i++) {
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
		printf("Chama falha a AddItems. Error code = %x\n", hr);
		exit(0);
	}

	// Server handle for the added item:
	for (i = 0; i < 3; i++) {
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
void RemoveGroup(IOPCServer* pIOPCServer, OPCHANDLE hServerGroup)
{
	// Remove the group:
	HRESULT hr = pIOPCServer->RemoveGroup(hServerGroup, FALSE);
	if (hr != S_OK) {
		if (hr == OPC_S_INUSE)
			printf("Falha ao remover grupo OPC. Ainda há alguma referência a ele.\n");
		else printf("Falha ao remover grupo OPC. Error code = %x\n", hr);
		exit(0);
	}
}

void wait(int time) {
	DWORD ticks1, ticks2;
	ticks1 = GetTickCount();
	ticks2 = GetTickCount();
	while ((ticks2 - ticks1) < time) {
		ticks2 = GetTickCount();
	}
}

////////////////////////////////////////////////////////////////////////
// Writes the value pointed by varValue to the item whose handle is 
// hServerItem, using the IOPCSyncIO interface obtained
// from the group interface pointed by pGroupIUnknown
//
void WriteItem(IUnknown* pGroupIUnknown, OPCHANDLE hServerItem, VARIANT *varValue)
{

	//get a pointer to the IOPCSyncIOInterface:
	IOPCSyncIO* pIOPCSyncIO;
	pGroupIUnknown->QueryInterface(__uuidof(pIOPCSyncIO), (void**)&pIOPCSyncIO);

	// write the item value to the device:
	HRESULT* pErrors = NULL; //to store error code(s)
	HRESULT hr = pIOPCSyncIO->Write(1, &hServerItem, varValue, &pErrors);
	_ASSERT(!hr);

	if (varValue->vt == 2) {
		printf("\nItem escrito: valor = %d\n", varValue->iVal);
	}
	else if (varValue->vt == 3) {
		printf("\nItem escrito: valor = %d\n", varValue->iVal);
	}
	else if (varValue->vt == 8) {
		printf("\nItem escrito: valor = %S\n", varValue->bstrVal);
	}
	//Release memeory allocated by the OPC server:
	CoTaskMemFree(pErrors);
	pErrors = NULL;

	// release the reference to the IOPCSyncIO interface:
	pIOPCSyncIO->Release();
}
