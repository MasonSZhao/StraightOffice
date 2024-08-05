// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#pragma once

#include <ole2.h>

namespace STRAIGHTOLE {
/**
 * @brief This class helps to initialize and uninitialize the COM library.
 */
class OleInit {
private:
    /**
     * @brief Initializes the COM library for use by the calling thread, sets the thread's concurrency model, and creates a new apartment for the thread if one is required.
     * @return Return code:
     *     - S_OK: Success.
     *     - S_FALSE: The COM library is already initialized on this thread.
     *     - RPC_E_CHANGED_MODE: A previous call to CoInitializeEx specified the concurrency model for this apartment as multithread apartment (MTA).\n
     *       This could also mean that a change from neutral threaded apartment to single threaded apartment occurred.
     * @details https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-coinitializeex
     */
    [[deprecated]] static HRESULT initComForAllThreads();

public:
    /**
     * @brief Initializes the COM library on the current thread, identifies the concurrency model as single-thread apartment (STA).
     * @return Return code:
     *     - S_OK: Success.
     *     - S_FALSE: The COM library is already initialized on this thread.
     *     - RPC_E_CHANGED_MODE: A previous call to CoInitializeEx specified the concurrency model for this apartment as multithread apartment (MTA).\n
     *       This could also mean that a change from neutral threaded apartment to single threaded apartment occurred.
     * @details https://learn.microsoft.com/en-us/windows/win32/api/objbase/nf-objbase-coinitialize
     */
    static HRESULT initComForThisThread();

    /**
     * @brief Closes the COM library on the current thread, unloads all DLLs loaded by the thread, frees any other resources that the thread maintains, and forces all RPC connections on the thread to close.
     * @details https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-couninitialize
     */
    static void uninitComForThisThread();
};

}
