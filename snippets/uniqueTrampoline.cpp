/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */

#pragma warning(push)
#pragma warning(disable : 4668 4820 4917)

#include <iostream>

#include <Windows.h>

#pragma warning(pop)

static /* thread_local */ int nFoo;

static int foobar()
{
    std::cout << "Hehe, this is " << nFoo << "\n";
    return nFoo + 1000;
}

static int uniqueAddress(void* pFunction, void** pUnique)
{
    static int nCounter = 0;

    static SYSTEM_INFO aSystemInfo;

    if (nCounter == 0)
    {
        GetSystemInfo(&aSystemInfo);
        std::cout << "Page size is " << aSystemInfo.dwPageSize << "\n";
    }

    static unsigned char* pPage
        = (unsigned char*)VirtualAlloc(NULL, aSystemInfo.dwPageSize, MEM_COMMIT, PAGE_READWRITE);
    if (pPage == NULL)
    {
        std::cerr << "VirtualAlloc failed\n";
        std::exit(1);
    }
    static unsigned char* pCode = pPage;

    if (pCode > pPage + aSystemInfo.dwPageSize - 100)
    {
        pPage = (unsigned char*)VirtualAlloc(NULL, aSystemInfo.dwPageSize, MEM_COMMIT,
                                             PAGE_READWRITE);
        if (pPage == NULL)
        {
            std::cerr << "VirtualAlloc failed\n";
            std::exit(1);
        }
        pCode = pPage;
    }

    *pUnique = pCode;

    DWORD nOldProtection;
    if (!VirtualProtect(pPage, aSystemInfo.dwPageSize, PAGE_READWRITE, &nOldProtection))
    {
        std::cerr << "VirtualProtect failed\n";
        std::exit(1);
    }

    uintptr_t nAddressOfFoo = (uintptr_t)&nFoo;
#ifndef _WIN64
    // mov dword ptr [nFoo], <nCounter>
    *pCode++ = 0xC7;
    *pCode++ = 0x05;

    *pCode++ = ((unsigned char*)&nAddressOfFoo)[0];
    *pCode++ = ((unsigned char*)&nAddressOfFoo)[1];
    *pCode++ = ((unsigned char*)&nAddressOfFoo)[2];
    *pCode++ = ((unsigned char*)&nAddressOfFoo)[3];
    *pCode++ = ((unsigned char*)&nCounter)[0];
    *pCode++ = ((unsigned char*)&nCounter)[1];
    *pCode++ = ((unsigned char*)&nCounter)[2];
    *pCode++ = ((unsigned char*)&nCounter)[3];

    // jmp <relative 32-bit offset>
    *pCode++ = 0xE9;
    unsigned char* p = pCode;
    pCode += 4;

    intptr_t nDiff = ((unsigned char*)pFunction - pCode);
    *p++ = ((unsigned char*)&nDiff)[0];
    *p++ = ((unsigned char*)&nDiff)[1];
    *p++ = ((unsigned char*)&nDiff)[2];
    *p++ = ((unsigned char*)&nDiff)[3];
#else
    // mov rax, [dword ptr nFoo] (or something like that)
    *pCode++ = 0x48;
    *pCode++ = 0xB8;
    *pCode++ = ((unsigned char*)&nAddressOfFoo)[0];
    *pCode++ = ((unsigned char*)&nAddressOfFoo)[1];
    *pCode++ = ((unsigned char*)&nAddressOfFoo)[2];
    *pCode++ = ((unsigned char*)&nAddressOfFoo)[3];
    *pCode++ = ((unsigned char*)&nAddressOfFoo)[4];
    *pCode++ = ((unsigned char*)&nAddressOfFoo)[5];
    *pCode++ = ((unsigned char*)&nAddressOfFoo)[6];
    *pCode++ = ((unsigned char*)&nAddressOfFoo)[7];

    // mov dword ptr [rax], <nCounter>
    *pCode++ = 0xC7;
    *pCode++ = 0x00;
    *pCode++ = ((unsigned char*)&nCounter)[0];
    *pCode++ = ((unsigned char*)&nCounter)[1];
    *pCode++ = ((unsigned char*)&nCounter)[2];
    *pCode++ = ((unsigned char*)&nCounter)[3];

    // mov rax, [dword ptr pFunction] (or something like that)
    *pCode++ = 0x48;
    *pCode++ = 0xB8;
    *pCode++ = ((unsigned char*)&pFunction)[0];
    *pCode++ = ((unsigned char*)&pFunction)[1];
    *pCode++ = ((unsigned char*)&pFunction)[2];
    *pCode++ = ((unsigned char*)&pFunction)[3];
    *pCode++ = ((unsigned char*)&pFunction)[4];
    *pCode++ = ((unsigned char*)&pFunction)[5];
    *pCode++ = ((unsigned char*)&pFunction)[6];
    *pCode++ = ((unsigned char*)&pFunction)[7];

    // jmp rax
    *pCode++ = 0xFF;
    *pCode++ = 0xE0;
#endif

    if (!VirtualProtect(pPage, aSystemInfo.dwPageSize, PAGE_EXECUTE, &nOldProtection))
    {
        std::cerr << "VirtualProtect failed\n";
        std::exit(1);
    }

    return nCounter++;
}

int main(int, char**)
{
    for (int i = 0; i < 1000; i++)
    {
        void* p;
        int nWhich = uniqueAddress(foobar, &p);
        std::cout << "Set it up at " << p << " with unique id of " << nWhich << "\n";
        int n = ((int (*)())p)();
        std::cout << "Got " << n << std::endl;
    }
    return 0;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */
