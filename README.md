# NotSoSimpleOPCClient
Repositório dedicado a hostear as alterações feitas ao código SimpleOPCClient para que realize escrita síncrona e leitura assíncrona de múltiplos itens

Simple OPC Client
 This is a modified version of the "Simple OPC Client" originally
 developed by Philippe Gras (CERN) for demonstrating the basic techniques
 involved in the development of an OPC DA client, and furtherly modified by UFMG's professor Luiz T. S. Mendes.

 The modifications made by prof. Luiz are the introduction of two C++ classes to allow the
 the client to ask for callback notifications from the OPC server, and
 the corresponding introduction of a message comsumption loop in the
 main program to allow the client to process those notifications. The
 C++ classes implement the OPC DA 1.0 IAdviseSink and the OPC DA 2.0
 IOPCDataCallback client interfaces, and in turn were adapted from the
 KEPWARE´s  OPC client sample code. A few wrapper functions to initiate
 and to cancel the notifications were also developed.

 Further to prof. Luiz's modifications, my modifications were:
 Removal of the C++ classes that implemented the OPC DA 1.0 IAdviseSink
 
 -Removal of the C++ class that implemented synchronous reading
 
 -Introduction of 6 new items, 3 of them to be read and 3 to be written on
 
 -Introduction of a C++ class that implements synchronous writing using the IOPCSyncIO interface
 
 -Initialization of the environment using multi thread apartments, instead of the original single thread apartments
 
 -Removal of the message comsumption loop, no longer necessary with the MTA newly introduced approach
 
 -Replacement of the C++ class AddTheItems by two similar classes, one for the reading itens and another for the writing ones, both capable of adding arrays of items instead of single items.

 -Changes to the AddTheGroup C++ class, for it to take the group name as a parameter instead of a fixed one and changing the reading interval to 500ms.
 
 -Creation of two groups - one for reading and another for writing -, both with multiple items.
 
 -Addition of case VT_DATE to the method VarToStr, with proper conversion between these types implemented as well.
 
 -Introduction of multithread behavior to the code, transfering all the OPC calls to the OPCThread method.
 
 -Introduction of mutexes and semaphores to deal with multiple access to variables
 
 -Integration of a socket server to the code, inside it's own thread, which runs the method SocketThread

 The original Simple OPC Client code can still be found (as of this date)
 in
        http://pgras.home.cern.ch/pgras/OPCClientTutorial/


Elise Guimaraes de Araujo - UFMG - 26/05/2016
elise.guimaraes at ufmg.br
