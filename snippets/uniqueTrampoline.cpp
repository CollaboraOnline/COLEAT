/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4; fill-column: 100 -*- */

#pragma warning(push)
#pragma warning(disable : 4668 4820 4917)

#include <cassert>
#include <iostream>

#include <Windows.h>

#pragma warning(pop)

#pragma runtime_checks("", off)
#pragma optimize("", off)

HRESULT __stdcall P0() { return 0; }

HRESULT __stdcall myP0() { return P0(); }

HRESULT __stdcall P1(LPVOID) { return 0; }

HRESULT __stdcall myP1(LPVOID arg0) { return P1(arg0); }

HRESULT __stdcall P2(LPVOID, LPVOID) { return 0; }

HRESULT __stdcall myP2(LPVOID arg0, LPVOID arg1) { return P2(arg0, arg1); }

HRESULT __stdcall P3(LPVOID, LPVOID, LPVOID) { return 0; }

HRESULT __stdcall myP3(LPVOID arg0, LPVOID arg1, LPVOID arg2) { return P3(arg0, arg1, arg2); }

HRESULT __stdcall P4(LPVOID, LPVOID, LPVOID, LPVOID) { return 0; }

HRESULT __stdcall myP4(LPVOID arg0, LPVOID arg1, LPVOID arg2, LPVOID arg3)
{
    return P4(arg0, arg1, arg2, arg3);
}

HRESULT __stdcall P5(LPVOID, LPVOID, LPVOID, LPVOID, LPVOID) { return 0; }

HRESULT __stdcall myP5(LPVOID arg0, LPVOID arg1, LPVOID arg2, LPVOID arg3, LPVOID arg4)
{
    return P5(arg0, arg1, arg2, arg3, arg4);
}

HRESULT __stdcall P6(LPVOID, LPVOID, LPVOID, LPVOID, LPVOID, LPVOID) { return 0; }

HRESULT __stdcall myP6(LPVOID arg0, LPVOID arg1, LPVOID arg2, LPVOID arg3, LPVOID arg4, LPVOID arg5)
{
    return P6(arg0, arg1, arg2, arg3, arg4, arg5);
}

HRESULT __stdcall P7(LPVOID, LPVOID, LPVOID, LPVOID, LPVOID, LPVOID, LPVOID) { return 0; }

HRESULT __stdcall myP7(LPVOID arg0, LPVOID arg1, LPVOID arg2, LPVOID arg3, LPVOID arg4, LPVOID arg5,
                       LPVOID arg6)
{
    return P7(arg0, arg1, arg2, arg3, arg4, arg5, arg6);
}

HRESULT __stdcall P8(LPVOID, LPVOID, LPVOID, LPVOID, LPVOID, LPVOID, LPVOID, LPVOID) { return 0; }

HRESULT __stdcall myP8(LPVOID arg0, LPVOID arg1, LPVOID arg2, LPVOID arg3, LPVOID arg4, LPVOID arg5,
                       LPVOID arg6, LPVOID arg7)
{
    return P8(arg0, arg1, arg2, arg3, arg4, arg5, arg6, arg7);
}

HRESULT __stdcall P9(LPVOID, LPVOID, LPVOID, LPVOID, LPVOID, LPVOID, LPVOID, LPVOID, LPVOID)
{
    return 0;
}

HRESULT __stdcall myP9(LPVOID arg0, LPVOID arg1, LPVOID arg2, LPVOID arg3, LPVOID arg4, LPVOID arg5,
                       LPVOID arg6, LPVOID arg7, LPVOID arg8)
{
    return P9(arg0, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
}

HRESULT __stdcall P11(LPVOID, LPVOID, LPVOID, LPVOID, LPVOID, LPVOID, LPVOID, LPVOID, LPVOID,
                      LPVOID, LPVOID)
{
    return 0;
}

HRESULT __stdcall myP11(LPVOID arg0, LPVOID arg1, LPVOID arg2, LPVOID arg3, LPVOID arg4,
                        LPVOID arg5, LPVOID arg6, LPVOID arg7, LPVOID arg8, LPVOID arg9,
                        LPVOID arg10)
{
    return P11(arg0, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10);
}

#pragma optimize("", on)
#pragma runtime_checks("", restore)

static int __stdcall foobar1(int nA)
{
#ifndef _WIN64
    // 9 = the length of the __stdcall epilogue
    unsigned char* pId = (unsigned char*)_ReturnAddress() + 9;
#else
    // 6 = the length of the epilogue
    unsigned char* pId = (unsigned char*)_ReturnAddress() + 6;
#endif
    int nId = pId[0] + 0x100 * pId[1] + 0x10000 * pId[2];

    std::cout << "foobar1 instance " << nId << " (" << nA << ")" << std::endl;

    return nId + 1000;
}

static int __stdcall foobar3(int nA, char* sB, int nC)
{
#ifndef _WIN64
    // 9 = the length of the __stdcall epilogue
    unsigned char* pId = (unsigned char*)_ReturnAddress() + 9;
#else
    // 6 = the length of the epilogue
    unsigned char* pId = (unsigned char*)_ReturnAddress() + 6;
#endif

    int nId = pId[0] + 0x100 * pId[1] + 0x10000 * pId[2];

    std::cout << "foobar3 instance " << nId << " (" << nA << ", " << sB << ", " << nC << ")"
              << std::endl;

    return nId + 1000;
}

static void* generateTrampoline(void* pFunction, uintptr_t nId, short nArguments)
{
    // We must allocate a fresh page for each trampoline because of multi-thread concerns: Otherwise
    // we would need to change the protection of the page back to RW for a moment when creating
    // another trampoline on it, and if another thread was just executing an existing trampoline,
    // that would be a problem.
    unsigned char* pPage = (unsigned char*)VirtualAlloc(NULL, 100, MEM_COMMIT, PAGE_READWRITE);
    if (pPage == NULL)
    {
        std::cerr << "VirtualAlloc failed\n";
        std::exit(1);
    }
    unsigned char* pCode = pPage;

#ifndef _WIN64
    // Normal __stdcall prologue

    // push ebp
    *pCode++ = 0x55;

    // mov evp, esp
    *pCode++ = 0x8B;
    *pCode++ = 0xEC;

    // sub esp, 64
    *pCode++ = 0x83;
    *pCode++ = 0xEC;
    *pCode++ = 0x40;

    // push ebx
    *pCode++ = 0x53;

    // push esi
    *pCode++ = 0x56;

    // push edi
    *pCode++ = 0x57;

    // Push our parameters
    for (short i = 0; i < nArguments; ++i)
    {
        if ((i % 3) == 0)
        {
            // mov eax, dword ptr arg[ebp]
            *pCode++ = 0x8B;
            *pCode++ = 0x45;
            *pCode++ = (unsigned char)(8 + (nArguments - i - 1) * 4);

            // push eax
            *pCode++ = 0x50;
        }
        else if ((i % 3) == 1)
        {
            // mov ecx, dword ptr arg[ebp]
            *pCode++ = 0x8B;
            *pCode++ = 0x4D;
            *pCode++ = (unsigned char)(8 + (nArguments - i - 1) * 4);

            // push ecx
            *pCode++ = 0x51;
        }
        else if ((i % 3) == 2)
        {
            // mov edx, dword ptr arg[ebp]
            *pCode++ = 0x8B;
            *pCode++ = 0x55;
            *pCode++ = (unsigned char)(8 + (nArguments - i - 1) * 4);

            // push edx
            *pCode++ = 0x52;
        }
        else
            abort();
    }

    // call <relative 32-bit offset>
    *pCode++ = 0xE8;
    intptr_t nDiff = ((unsigned char*)pFunction - pCode - 4);
    *pCode++ = ((unsigned char*)&nDiff)[0];
    *pCode++ = ((unsigned char*)&nDiff)[1];
    *pCode++ = ((unsigned char*)&nDiff)[2];
    *pCode++ = ((unsigned char*)&nDiff)[3];

    // Normal __stdcall epilogue

    // pop edi
    *pCode++ = 0x5F;

    // pop esi
    *pCode++ = 0x5E;

    // pop ebx
    *pCode++ = 0x5B;

    // mov esp, ebp
    *pCode++ = 0x8B;
    *pCode++ = 0xE5;

    // pop ebp
    *pCode++ = 0x5D;

    // ret <nArguments*4>
    *pCode++ = 0xC2;
    short n = nArguments * 4;
    *pCode++ = ((unsigned char*)&n)[0];
    *pCode++ = ((unsigned char*)&n)[1];

    // the unique id is stored after the ret <n>
    *pCode++ = ((unsigned char*)&nId)[0];
    *pCode++ = ((unsigned char*)&nId)[1];
    *pCode++ = ((unsigned char*)&nId)[2];
    *pCode++ = ((unsigned char*)&nId)[3];
#else
    // Normal prologue

    if (nArguments > 3)
    {
        // mov qword ptr [rsp+32], r9
        *pCode++ = 0x4C;
        *pCode++ = 0x89;
        *pCode++ = 0x4C;
        *pCode++ = 0x24;
        *pCode++ = 0x20;
    }

    if (nArguments > 2)
    {
        // mov qword ptr [rsp+24], r8
        *pCode++ = 0x4C;
        *pCode++ = 0x89;
        *pCode++ = 0x44;
        *pCode++ = 0x24;
        *pCode++ = 0x18;
    }

    if (nArguments > 1)
    {
        // mov qword ptr [rsp+16], rdx
        *pCode++ = 0x48;
        *pCode++ = 0x89;
        *pCode++ = 0x54;
        *pCode++ = 0x24;
        *pCode++ = 0x10;
    }

    if (nArguments > 0)
    {
        // mov qword ptr [rsp+8], rcx
        *pCode++ = 0x48;
        *pCode++ = 0x89;
        *pCode++ = 0x4C;
        *pCode++ = 0x24;
        *pCode++ = 0x08;
    }

    // push rbp
    *pCode++ = 0x40;
    *pCode++ = 0x55;

    // sub rsp, <x>
    *pCode++ = 0x48;
    *pCode++ = 0x83;
    *pCode++ = 0xEC;
    if (nArguments <= 4)
        *pCode++ = 0x60;
    else
    {
        assert(0x70 + (nArguments - 5) / 2 * 0x10 <= 0xFF);
        *pCode++ = (unsigned char)(0x70 + (nArguments - 5) / 2 * 0x10);
    }

    // lea rbp, qword ptr [rsp+<x>]
    *pCode++ = 0x48;
    *pCode++ = 0x8D;
    *pCode++ = 0x6C;
    *pCode++ = 0x24;
    if (nArguments <= 4)
        *pCode++ = 0x20;
    else
    {
        assert(0x30 + (nArguments - 5) / 2 * 0x10 <= 0xFF);
        *pCode++ = (unsigned char)(0x30 + (nArguments - 5) / 2 * 0x10);
    }

    // Parameters
    for (short i = 0; i < nArguments; ++i)
    {
        if (i == 0)
        {
            // mov rcx, qword ptr arg0[rbp]
            *pCode++ = 0x48;
            *pCode++ = 0x8B;
            *pCode++ = 0x4D;
            *pCode++ = 0x50;
        }
        else if (i == 1)
        {
            // mov rdx, qword ptr arg1[rbp]
            *pCode++ = 0x48;
            *pCode++ = 0x8B;
            *pCode++ = 0x55;
            *pCode++ = 0x58;
        }
        else if (i == 2)
        {
            // mov r8, qword ptr arg2[rbp]
            *pCode++ = 0x4C;
            *pCode++ = 0x8B;
            *pCode++ = 0x45;
            *pCode++ = 0x60;
        }
        else if (i == 3)
        {
            // mov r9, qword ptr arg2[rbp]
            *pCode++ = 0x4C;
            *pCode++ = 0x8B;
            *pCode++ = 0x4D;
            *pCode++ = 0x68;
        }
        else
        {
            if (i <= 5)
            {
                // mov rax, qword ptr (112+(i-4)*8)[rbp]
                *pCode++ = 0x48;
                *pCode++ = 0x8B;
                *pCode++ = 0x45;
                *pCode++ = (unsigned char)(112 + (i - 4) * 8);
            }
            else
            {
                // mov rax, qword ptr (112+(i-4)*8)[rbp]
                *pCode++ = 0x48;
                *pCode++ = 0x8B;
                *pCode++ = 0x85;
                int n = 112 + (i - 4) * 8;
                *pCode++ = ((unsigned char*)&n)[0];
                *pCode++ = ((unsigned char*)&n)[1];
                *pCode++ = ((unsigned char*)&n)[2];
                *pCode++ = ((unsigned char*)&n)[3];
            }

            // mov qword ptr [rsp+32+(i-4)*8], rax
            *pCode++ = 0x48;
            *pCode++ = 0x89;
            *pCode++ = 0x44;
            *pCode++ = 0x24;
            *pCode++ = (unsigned char)(32 + (i - 4) * 8);
        }
    }

    // mov rax, qword ptr pFunction (or something like that)
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

    // call rax
    *pCode++ = 0xFF;
    *pCode++ = 0xD0;

    // Normal epilogue

    // lea rsp, qword ptr [rbp+64]
    *pCode++ = 0x48;
    *pCode++ = 0x8D;
    *pCode++ = 0x65;
    *pCode++ = 0x40;

    // pop rbp
    *pCode++ = 0x5D;

    // ret
    *pCode++ = 0xC3;

    // the unique number is stored after the ret
    *pCode++ = ((unsigned char*)&nId)[0];
    *pCode++ = ((unsigned char*)&nId)[1];
    *pCode++ = ((unsigned char*)&nId)[2];
    *pCode++ = ((unsigned char*)&nId)[3];
#endif
    DWORD nOldProtection;
    if (!VirtualProtect(pPage, 100, PAGE_EXECUTE, &nOldProtection))
    {
        std::cerr << "VirtualProtect failed\n";
        std::exit(1);
    }

    return pPage;
}

int main(int, char**)
{
    for (uintptr_t i = 0; i < 100; i++)
    {
        int n;

        void* p1 = generateTrampoline(foobar1, i, 1);
        std::cout << "Set foobar1 up at " << p1 << " with unique id of " << i << "\n";
        n = ((int(__stdcall*)(int))p1)((int)i);
        std::cout << "Got " << n << std::endl;

        void* p3 = generateTrampoline(foobar3, i * 3, 3);
        std::cout << "Set foobar3 up at " << p3 << " with unique id of " << i * 3 << "\n";
        n = ((int(__stdcall*)(int, char*, int))p3)((int)i, "hai", (int)i * 2);
        std::cout << "Got " << n << std::endl;
    }
    return 0;
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */
