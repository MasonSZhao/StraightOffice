// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#pragma once

#include "StraightWordDispatchDriver.h"
#include "StraightWordApplication.h"

namespace STRAIGHTWORD {
/**
 * @brief Represents Word program.
 */
class WordProgram {
private:
    /**
     * @brief Return the clsid of Excel.
     * @param des The destination to return to.
     * @return Return code:
     *     - S_OK: The CLSID was retrieved successfully.
     *     - CO_E_CLASSSTRING: The registered CLSID for the ProgID is invalid.
     *     - REGDB_E_WRITEREGDB: An error occurred writing the CLSID to the registry.
     * @details https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-clsidfromprogid
     */
    static HRESULT getCLSIDFromProgID(CLSID& des);

public:
    /**
     * @brief Starts an Excel application.
     * @param des The destination to start to.
     * @return Return code:
     *     - S_OK: The CLSID was retrieved successfully.
     *     - CO_E_CLASSSTRING: The registered CLSID for the ProgID is invalid.
     *     - REGDB_E_WRITEREGDB: An error occurred writing the CLSID to the registry.
     *     - S_OK: An instance of the specified object class was successfully created.
     *     - REGDB_E_CLASSNOTREG: A specified class is not registered in the registration database. Also can indicate that the type of server you requested in the CLSCTX enumeration is not registered or the values for the server types in the registry are corrupt.
     *     - CLASS_E_NOAGGREGATION: This class cannot be created as part of an aggregate.
     *     - E_NOINTERFACE: The specified class does not implement the requested interface, or the controlling IUnknown does not expose the requested interface.
     *     - E_POINTER: The ppv parameter is NULL.
     * @details
     *     https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-clsidfromprogid
     *
     *     https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cocreateinstance
     */
    static HRESULT start(WordApplication& des);

    /**
     * @brief Connects to a running Excel application.
     * @param des The destination to connect to.
     * @return Return code:
     *     - S_OK: The CLSID was retrieved successfully.
     *     - CO_E_CLASSSTRING: The registered CLSID for the ProgID is invalid.
     *     - REGDB_E_WRITEREGDB: An error occurred writing the CLSID to the registry.
     *     - S_OK: Success.
     *     - S_OK: The interface is supported.
     *     - E_NOINTERFACE: The interface is NOT supported.
     *     - E_POINTER: The ppvObject (the address) is nullptr.
     * @details
     *     https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-clsidfromprogid
     *
     *     https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-getactiveobject
     *
     *     https://learn.microsoft.com/en-us/windows/win32/api/unknwn/nf-unknwn-iunknown-queryinterface(refiid_void)
     */
    static HRESULT conn(WordApplication& des);
};

}
