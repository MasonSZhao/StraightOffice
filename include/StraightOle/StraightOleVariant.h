// Copyright (c) 2024-present, Shuai (Mason) Zhao
// Distributed under the MIT License (http://opensource.org/licenses/MIT)

#pragma once

#include <ole2.h>
#include <string>

namespace STRAIGHTOLE {
/**
 * @brief Encapsulates the VARIANT data type.
 * @details https://learn.microsoft.com/en-us/cpp/mfc/reference/colevariant-class?view=msvc-170#clear
 */
struct OleVariant : public tagVARIANT {
private:
    /**
     * @brief Call this function to attach the given VARIANT object to the current COleVariant object.
     * @param variant An existing VARIANT object to be attached to the current COleVariant object.
     */
    [[deprecated("")]] void Attach(VARIANT& variant);

    /**
     * @brief Converts the type of variant value in this COleVariant object.
     * @param vt The VARENUM for this COleVariant object.
     * @param variant A pointer to the VARIANT object to be converted. If this value is NULL, this OleVariant object is used as the source for the conversion.
     */
    [[deprecated("")]] void ChangeType(VARENUM vt, VARIANT* variant);

    /**
     * @brief Clears the VARIANT.
     */
    [[deprecated("")]] void Clear();

    /**
     * @brief Detaches the underlying VARIANT object from this COleVariant object.
     * @return The resulting VARIANT structure.
     */
    [[deprecated("")]] VARIANT Detach();

protected:
    /**
     * @brief Defines the destructor which clears the inner VARIANT.
     */
    ~OleVariant();
};

/**
 * @brief Encapsulates the bool data type.
 */
struct OleBool : public OleVariant {
    /**
     * @brief The constructor.
     * @param val The value to initialize the inner VARIANT to.
     */
    OleBool(bool val);
};

/**
 * @brief Encapsulates the char data type.
 */
class OleChar : public OleVariant {
public:
    /**
     * @brief The constructor.
     * @param val The value to initialize the inner VARIANT to.
     */
    OleChar(char val);
};

/**
 * @brief Encapsulates the unsigned char data type.
 */
class OleUChar : public OleVariant {
public:
    /**
     * @brief The constructor.
     * @param val The value to initialize the inner VARIANT to.
     */
    OleUChar(unsigned char val);
};

/**
 * @brief Encapsulates the short data type.
 */
class OleShrt : public OleVariant {
public:
    /**
     * @brief The constructor.
     * @param val The value to initialize the inner VARIANT to.
     */
    OleShrt(short val);
};

/**
 * @brief Encapsulates the unsigned short data type.
 */
class OleUShrt : public OleVariant {
public:
    /**
     * @brief The constructor.
     * @param val The value to initialize the inner VARIANT to.
     */
    OleUShrt(unsigned short val);
};

/**
 * @brief Encapsulates the int data type.
 */
class OleInt : public OleVariant {
public:
    /**
     * @brief The constructor.
     * @param val The value to initialize the inner VARIANT to.
     */
    OleInt(int val);
};

/**
 * @brief Encapsulates the unsigned int data type.
 */
class OleUInt : public OleVariant {
public:
    /**
     * @brief The constructor.
     * @param val The value to initialize the inner VARIANT to.
     */
    OleUInt(unsigned int val);
};

/**
 * @brief Encapsulates the long data type.
 */
class OleLong : public OleVariant {
public:
    /**
     * @brief The constructor.
     * @param val The value to initialize the inner VARIANT to.
     */
    OleLong(long val);
};

/**
 * @brief Encapsulates the unsigned long data type.
 */
class OleULong : public OleVariant {
public:
    /**
     * @brief The constructor.
     * @param val The value to initialize the inner VARIANT to.
     */
    OleULong(unsigned long val);
};

/**
 * @brief Encapsulates the long long data type.
 */
class OleLLong : public OleVariant {
public:
    /**
     * @brief The constructor.
     * @param val The value to initialize the inner VARIANT to.
     */
    OleLLong(long long val);
};

/**
 * @brief Encapsulates the unsigned long long data type.
 */
class OleULLong : public OleVariant {
public:
    /**
     * @brief The constructor.
     * @param val The value to initialize the inner VARIANT to.
     */
    OleULLong(unsigned long long val);
};

/**
 * @brief Encapsulates the float data type.
 */
class OleFlt : public OleVariant {
public:
    /**
     * @brief The constructor.
     * @param val The value to initialize the inner VARIANT to.
     */
    OleFlt(float val);
};

/**
 * @brief Encapsulates the double data type.
 */
class OleDbl : public OleVariant {
public:
    /**
     * @brief The constructor.
     * @param val The value to initialize the inner VARIANT to.
     */
    OleDbl(double val);
};

/**
 * @brief Encapsulates the string data type.
 */
struct OleStrA : public OleVariant {
public:
    /**
     * @brief The constructor.
     * @param val The value to initialize the inner VARIANT to.
     */
    OleStrA(const std::string& val);
};

/**
 * @brief Encapsulates the wstring data type.
 */
struct OleStrW : public OleVariant {
public:
    /**
     * @brief The constructor.
     * @param val The value to initialize the inner VARIANT to.
     */
    OleStrW(const std::wstring& val);
};

}
