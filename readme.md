# COM local server exe template

Minimalistic solution for the classic COM Local Server (exe) implementing IUnknown and CLI client written without ATL and only with Win32 API

For both: x64 and x86.

Inspired by:
 - https://www.codeproject.com/Articles/8679/Building-a-LOCAL-COM-Server-and-Client-A-Step-by-S
 - https://www.akella.org/homepages/mani/documents/secdocs/ASP.net/ComAndDotNetInteroperability.pdf
 - https://github.com/Apress/com-.net-interoperability

Usage:
1. register COM server with `com-server.exe /r`
2. run `test-client.exe`, this will lead to:
   * com-server.exe will to be started by SCM as a separate process
   * com-server.exe will be able service incoming COM calls
   * with each new run of test-client.exe the COM coclass will be created, executed its methods, and destroyed
   * the class factory, on the contrary, will remain alive, so to shutdown the process, you would need to:
3. `taskkill /f /im com-server.exe`
4. unregister COM server with `com-server.exe /u`
