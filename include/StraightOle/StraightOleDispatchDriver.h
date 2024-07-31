// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#pragma once

#include <ole2.h>
#include <string>

namespace STRAIGHTOLE {
/**
 * @brief This driver class makes it possible for the C++ application to manipulate objects implemented in another application.
 * @details
 *     https://learn.microsoft.com/en-us/cpp/mfc/reference/coledispatchdriver-class?view=msvc-170#invokehelper
 *
 *     https://learn.microsoft.com/en-us/cpp/mfc/automation-clients?view=msvc-170
 *
 *     https://learn.microsoft.com/en-us/cpp/mfc/automation-servers?view=msvc-170
 */
class OleDispatchDriver {
private:
    /**
     * @brief Constructs an OleDispatchDriver object with an IDispatch pointer.
     */
    [[deprecated("")]] OleDispatchDriver(IDispatch* dispatch, bool autoRelease = true);

    /**
     * @brief Attaches an IDispatch pointer to the OleDispatchDriver object.
     */
    [[deprecated("")]] void AttachDispatch(IDispatch* dispatch, bool autoRelease = true);

    /**
     * @brief Creates an IDispatch interface object and attaches it to the OleDispatchDriver object.
     * @param clsid Class ID of the IDispatch connection object to be created.
     * @return Nonzero on success; otherwise 0.
     */
    [[deprecated("")]] bool CreateDispatch(CLSID clsid);

    /**
     * @brief Looks up a Class ID in the registry, and Creates an IDispatch interface object with the given ProgID and attaches it to the OleDispatchDriver object.
     * @param progID A pointer to the ProgID whose CLSID is requested.
     * @return Nonzero on success; otherwise 0.
     */
    [[deprecated("")]] bool CreateDispatch(const wchar_t* progID);

    /**
     * @brief Detaches the current IDispatch connection from this object.
     * @return A pointer to the previously attached OLE IDispatch object.
     */
    [[deprecated("")]] IDispatch* DetachDispatch();

    /**
     * @brief Accesses the underlying IDispatch pointer of the COleDispatchDriver object.
     * @return The underlying IDispatch.
     */
    [[deprecated("")]] IDispatch* get_IDispatch();

    /**
     * @brief Copies the source value into the OleDispatchDriver object.
     * @param src A pointer to an existing source OleDispatchDriver object.
     * @return The const reference.
     */
    [[deprecated("")]] const OleDispatchDriver& operator=(const OleDispatchDriver& src);

protected:
    /**
     * @brief Specifies whether to release the IDispatch during ReleaseDispatch or object destruction.
     */
    bool _autoRelease { true };

    /**
     * @brief Indicates the pointer to the IDispatch interface attached to this COleDispatchDriver.
     */
    IDispatch* _dispatch { NULL };

    /**
     * @brief The inner invoke function.
     * @param autoType One of DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT and DISPATCH_PROPERTYPUTREF, see details.
     * @param pvResult The pointer of the VARIANT result.
     * @param dispID The DISP ID.
     * @param ... The Arguments.
     * @return Return code:
     *     - E_ABORT: The inner IDispatch is NULL, operation aborted.
     *     - S_OK: Success.
     *     - DISP_E_BADPARAMCOUNT: The number of elements provided to DISPPARAMS is different from the number of arguments accepted by the method or property.
     *     - DISP_E_BADVARTYPE: One of the arguments in DISPPARAMS is not a valid variant type.
     *     - DISP_E_EXCEPTION: The application needs to raise an exception. In this case, the structure passed in pexcepinfo should be filled in.
     *     - DISP_E_MEMBERNOTFOUND: The requested member does not exist.
     *     - DISP_E_NONAMEDARGS: This implementation of IDispatch does not support named arguments.
     *     - DISP_E_OVERFLOW: One of the arguments in DISPPARAMS could not be coerced to the specified type.
     *     - DISP_E_PARAMNOTFOUND: One of the parameter IDs does not correspond to a parameter on the method. In this case, puArgErr is set to the first argument that contains the error.
     *     - DISP_E_TYPEMISMATCH: One or more of the arguments could not be coerced. The index of the first parameter with the incorrect type within rgvarg is returned in puArgErr.
     *     - DISP_E_UNKNOWNINTERFACE: The interface identifier passed in riid is not IID_NULL.
     *     - DISP_E_UNKNOWNLCID: The member being invoked interprets string arguments according to the LCID, and the LCID is not recognized. If the LCID is not needed to interpret arguments, this error should not be returned
     *     - DISP_E_PARAMNOTOPTIONAL: A required parameter was omitted.
     * @details
     *     - DISPATCH_METHOD: The member is invoked as a method. If a property has the same name, both this and the DISPATCH_PROPERTYGET flag can be set.
     *     - DISPATCH_PROPERTYGET: The member is retrieved as a property or data member.
     *     - DISPATCH_PROPERTYPUT: The member is changed as a property or data member.
     *     - DISPATCH_PROPERTYPUTREF: The member is changed by a reference assignment, rather than a value assignment. This flag is valid only when the property accepts a reference to an object.
     *
     *     https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-invoke
     */
    HRESULT invokeNameId(int autoType, VARIANT* pvResult, DISPID dispID, int cArgs...);

    /**
     * @brief The inner invoke function.
     * @param autoType One of DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT and DISPATCH_PROPERTYPUTREF, see details.
     * @param pvResult The pointer of the VARIANT result.
     * @param ptName The DISP Name.
     * @param ... The Arguments.
     * @return Return code:
     *     - E_ABORT: The inner IDispatch is NULL, operation aborted.
     *     - E_OUTOFMEMORY: Out of memory.
     *     - DISP_E_UNKNOWNNAME: One or more of the specified names were not known. The returned array of DISPIDs contains DISPID_UNKNOWN for each entry that corresponds to an unknown name.
     *     - DISP_E_UNKNOWNLCID: The member being invoked interprets string arguments according to the LCID, and the LCID is not recognized. If the LCID is not needed to interpret arguments, this error should not be returned
     *     - S_OK: Success.
     *     - DISP_E_BADPARAMCOUNT: The number of elements provided to DISPPARAMS is different from the number of arguments accepted by the method or property.
     *     - DISP_E_BADVARTYPE: One of the arguments in DISPPARAMS is not a valid variant type.
     *     - DISP_E_EXCEPTION: The application needs to raise an exception. In this case, the structure passed in pexcepinfo should be filled in.
     *     - DISP_E_MEMBERNOTFOUND: The requested member does not exist.
     *     - DISP_E_NONAMEDARGS: This implementation of IDispatch does not support named arguments.
     *     - DISP_E_OVERFLOW: One of the arguments in DISPPARAMS could not be coerced to the specified type.
     *     - DISP_E_PARAMNOTFOUND: One of the parameter IDs does not correspond to a parameter on the method. In this case, puArgErr is set to the first argument that contains the error.
     *     - DISP_E_TYPEMISMATCH: One or more of the arguments could not be coerced. The index of the first parameter with the incorrect type within rgvarg is returned in puArgErr.
     *     - DISP_E_UNKNOWNINTERFACE: The interface identifier passed in riid is not IID_NULL.
     *     - DISP_E_PARAMNOTOPTIONAL: A required parameter was omitted.
     * @details
     *     - DISPATCH_METHOD: The member is invoked as a method. If a property has the same name, both this and the DISPATCH_PROPERTYGET flag can be set.
     *     - DISPATCH_PROPERTYGET: The member is retrieved as a property or data member.
     *     - DISPATCH_PROPERTYPUT: The member is changed as a property or data member.
     *     - DISPATCH_PROPERTYPUTREF: The member is changed by a reference assignment, rather than a value assignment. This flag is valid only when the property accepts a reference to an object.
     *
     *     https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-getidsofnames
     *
     *     https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-invoke
     */
    HRESULT invokeNameStr(int autoType, VARIANT* pvResult, LPOLESTR ptName, int cArgs...);

    /**
     * @brief The inner string stores the ostream content which is used by operator << method.
     */
    std::string _ostreamStr { "" };

public:
    /**
     * @brief Deletes the copy constructor.
     * @param The object to copy.
     */
    OleDispatchDriver(const OleDispatchDriver& src) = delete;

    /**
     * @brief Defines the default constructor as the copy constructor is defined.
     */
    OleDispatchDriver();

    /**
     * @brief Defines the destructor which releases the inner IDispatch if set.
     */
    ~OleDispatchDriver();

    /**
     * @brief The operator << method.
     * @param os The ostream.
     * @return The ostream.
     */
    virtual std::ostream& operator<<(std::ostream& os) const;

    /**
     * @brief Releases the IDispatch connection.
     * @details If auto release has been set for this connection, this function calls IDispatch::Release before releasing the interface.
     */
    void ReleaseDispatch();
};

}
