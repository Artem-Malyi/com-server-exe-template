//
// locks.cpp
// This file contains the COM server locking functions.
//
#include "..\common\pch.h"
#include "locks.h"

#define DEBUG_LOGGER_ENABLED
#define FILE_LOGGER_ENABLED
#define LOG_PREFIX "[ADDOBJ-LOCKSRV]"
#include "logger.h"

extern long g_nComObjectsInUse;

// Called by coclass ctor and LockServer(TRUE)
void Lock()
{
    InterlockedIncrement(&g_nComObjectsInUse);
}

// Called by coclass dtor and LockServer(FALSE)
void UnLock()
{
    long nComObjectsInUse = InterlockedDecrement(&g_nComObjectsInUse);

    if (0 == nComObjectsInUse) {
        LOG("PostQuitMessage(EXIT_SUCCESS)");
        PostQuitMessage(EXIT_SUCCESS);
    }
}