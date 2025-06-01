#ifndef SIMNRA_H
#define SIMNRA_H

#include <windows.h>
#include <comdef.h>
#include <string>
#include <map>
#include <iostream>
#include <vector>
#include <chrono>

/**
 * @brief Creates a VARIANT representing an integer value.
 * @param val The integer value.
 * @return A VARIANT with the integer value.
 */
VARIANT MakeIntArg(int val);

/**
 * @brief Creates a VARIANT representing a double value.
 * @param val The double value.
 * @return A VARIANT with the double value.
 */
VARIANT MakeDoubleArg(double val);

/**
 * @brief Creates a VARIANT representing a wide string.
 * @param val The wide string.
 * @return A VARIANT with the BSTR value.
 */
VARIANT MakeStringArg(const std::wstring& val);

/**
 * @brief Checks HRESULT and throws an error if failed.
 * @param hr The HRESULT value.
 * @param msg Custom error message.
 */
void check_hresult(HRESULT hr, const std::string& msg) ;

/**
 * @brief Gets a property value from a COM object.
 * @param disp Pointer to IDispatch.
 * @param propertyName The property name.
 * @param args Arguments required to access the property.
 * @return The property value as a VARIANT.
 */
VARIANT GetPropertyValue(IDispatch* disp, const wchar_t* propertyName, const std::vector<VARIANT>& args);

/**
 * @brief Sets a property value on a COM object.
 * @param disp Pointer to IDispatch.
 * @param propertyName The property name.
 * @param args Arguments including the value to set.
 */
void SetPropertyValue(IDispatch* disp, const wchar_t* propertyName, const std::vector<VARIANT>& args);

/**
 * @brief Creates a COM object using its ProgID.
 * @param progID The programmatic identifier.
 * @return Pointer to IDispatch interface.
 */
IDispatch* CreateDispatch(const wchar_t* progID);

/**
 * @brief Invokes a method on a COM object that does not return a value.
 * @param disp Pointer to IDispatch.
 * @param method Method name.
 * @param args Optional arguments to pass.
 */
void InvokeVoidMethod(IDispatch* disp, const wchar_t* method, const std::vector<VARIANT>& args = {});

/**
 * @brief Converts a VARIANT to a specified type.
 * @tparam T Target C++ type.
 * @param var The VARIANT to convert.
 * @return The converted value.
 */
template<typename T>
T variantTo(const VARIANT& var) {
    if constexpr (std::is_same_v<T, int>) {
        if (var.vt == VT_I4) return var.lVal;
        throw std::runtime_error("Variant is not an int (VT_I4)");
    }
    else if constexpr (std::is_same_v<T, bool>) {
        if (var.vt == VT_BOOL) return var.boolVal == VARIANT_TRUE;
        throw std::runtime_error("Variant is not a bool (VT_BOOL)");
    }
    else if constexpr (std::is_same_v<T, double>) {
        if (var.vt == VT_R8) return var.dblVal;
        if (var.vt == VT_R4) return static_cast<double>(var.fltVal);
        throw std::runtime_error("Variant is not a double (VT_R8 or VT_R4)");
    }
    else if constexpr (std::is_same_v<T, std::wstring>) {
        if (var.vt == VT_BSTR) return std::wstring(var.bstrVal);
        throw std::runtime_error("Variant is not a BSTR (wstring)");
    }
    else if constexpr (std::is_same_v<T, std::string>) {
        if (var.vt == VT_BSTR) return std::string((const char*)_bstr_t(var.bstrVal));
        throw std::runtime_error("Variant is not a BSTR (string)");
    }
    else {
        static_assert(!std::is_same_v<T, T>, "Unsupported type in VariantTo<T>()");
    }
}

/**
 * @brief Invokes a method on a COM object and returns a typed result.
 * @tparam T Return type.
 * @param disp Pointer to IDispatch.
 * @param methodName Method name to invoke.
 * @param args Optional arguments to pass.
 * @return The result of the method invocation.
 */
template<typename T>
T InvokeMethod(IDispatch* disp, const wchar_t* methodName, const std::vector<VARIANT>& args = {}) {
    if (!disp) throw std::runtime_error("IDispatch is null");

    // Get DISPID
    DISPID dispid;
    LPOLESTR nameCopy = const_cast<LPOLESTR>(methodName);
    HRESULT hr = disp->GetIDsOfNames(IID_NULL, &nameCopy, 1, LOCALE_USER_DEFAULT, &dispid);
    check_hresult(hr, "GetIDsOfNames failed");

    // Copy and reverse arguments
    std::vector<VARIANT> reversedArgs(args.rbegin(), args.rend());

    DISPPARAMS dp;
    dp.cArgs = static_cast<UINT>(reversedArgs.size());
    dp.rgvarg = const_cast<VARIANT*>(reversedArgs.data());
    dp.cNamedArgs = 0;
    dp.rgdispidNamedArgs = nullptr;

    VARIANT result;
    VariantInit(&result);

    hr = disp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, nullptr, nullptr);
    check_hresult(hr, "Invoke failed");

    // Extract result
    T value{};
    if constexpr (std::is_same_v<T, double>) {
        if (result.vt == VT_R8)
            value = result.dblVal;
        else if (result.vt == VT_I4)
            value = static_cast<double>(result.intVal);
        else
            throw std::runtime_error("Unexpected return type for double");
    } else if constexpr (std::is_same_v<T, int>) {
        if (result.vt == VT_I4)
            value = result.intVal;
        else
            throw std::runtime_error("Unexpected return type for int");
    } else if constexpr (std::is_same_v<T, std::wstring>) {
        if (result.vt == VT_BSTR)
            value = std::wstring(result.bstrVal, SysStringLen(result.bstrVal));
        else
            throw std::runtime_error("Unexpected return type for string");
    }

    VariantClear(&result);
    return value;
}


/**
 * @brief Wrapper class for SIMNRA COM interface automation.
 */
class SIMNRA {
private:
    IDispatch* m_App = nullptr;
    IDispatch* m_Setup = nullptr;
    IDispatch* m_Target = nullptr;
    IDispatch* m_Calc = nullptr;
    IDispatch* m_Fit = nullptr;
    IDispatch* m_Projectile = nullptr;
    IDispatch* m_Spectrum = nullptr;
    IDispatch* m_Stopping = nullptr;
    IDispatch* m_PIGE = nullptr;
    IDispatch* m_CrossSec = nullptr;

public:

    /**
     * @brief Constructor that initializes COM and connects to SIMNRA components.
     */
    SIMNRA() {
        HRESULT hr = CoInitialize(NULL);
        if (FAILED(hr)) {
            throw std::runtime_error("Failed to initialize COM");
        }

        m_App = CreateDispatch(L"SIMNRA.app");
        m_Setup = CreateDispatch(L"SIMNRA.setup");
        m_Target = CreateDispatch(L"SIMNRA.target");
        m_Calc = CreateDispatch(L"SIMNRA.calc");
        m_Fit = CreateDispatch(L"SIMNRA.fit");
        m_Projectile = CreateDispatch(L"SIMNRA.projectile");
        m_Spectrum = CreateDispatch(L"SIMNRA.spectrum");
        m_Stopping = CreateDispatch(L"SIMNRA.stopping");
        m_PIGE = CreateDispatch(L"SIMNRA.pige");
        m_CrossSec = CreateDispatch(L"SIMNRA.crosssec");
    }

    /**
     * @brief Destructor that releases COM objects and uninitializes COM.
     */
    ~SIMNRA() {
        if (m_App) m_App->Release();
        if (m_Setup) m_Setup->Release();
        if (m_Target) m_Target->Release();
        if (m_Calc) m_Calc->Release();
        if (m_Fit) m_Fit->Release();
        if (m_Projectile) m_Projectile->Release();
        if (m_Spectrum) m_Spectrum->Release();
        if (m_Stopping) m_Stopping->Release();
        if (m_PIGE) m_PIGE->Release();
        if (m_CrossSec) m_CrossSec->Release();
        CoUninitialize();
    }

    /// @name SIMNRA.App methods
    /// @{   
    
    /**
     * @brief Gets the last message from SIMNRA.
     * @return The last message as a string.
     */
    std::wstring getLastMessage(){
        return variantTo<std::wstring>(GetPropertyValue(m_App, L"LastMessage", {}));
    };

    /**
     * @brief Opens a file in SIMNRA.
     * @param filename The file path.
     * @param type File type (default -1).
     */
    void open(const wchar_t* filename, int type=-1) {
        InvokeVoidMethod(m_App, L"Open", {MakeStringArg(filename), MakeIntArg(type)});
    }; 

    /**
     * @brief Calculates the spectrum using the full method.
     */
    void calculateSpectrum() {
        InvokeVoidMethod(m_App, L"CalculateSpectrum", {});
    };

    /**
     * @brief Calculates the spectrum using the fast method.
     */
    void calculateSpectrumFast(){
        InvokeVoidMethod(m_App, L"CalculateSpectrumFast", {});
    };


    /// @}

    /// @name SIMNRA.Target methods
    /// @{

    /**
     * @brief Gets the number of layers in the target.
     * @return Number of layers.
     */
    int getNumberOfLayers(){
        return variantTo<int>(GetPropertyValue(m_Target, L"NumberOfLayers", {}));
    };

    /**
     * @brief Sets the number of layers in the target.
     * @param numLayers The desired number of layers.
     */
    void setNumberOfLayers(int numLayers){
        SetPropertyValue(m_Target, L"NumberOfLayers", {MakeIntArg(numLayers)});
    };

    /**
     * @brief Gets the thickness of a specific layer.
     * @param layerIndex Index of the layer (default 1).
     * @return Thickness in appropriate units.
     */
    double getLayerThickness( int layerIndex=1) {
        return variantTo<double>(GetPropertyValue(m_Target, L"LayerThickness", {MakeIntArg(layerIndex),} ));
    }; 

    /**
     * @brief Sets the thickness of a specific layer.
     * @param layerIndex Index of the layer.
     * @param thick Thickness to set.
     */
    void setLayerThickness( int layerIndex, double thick ) {
        SetPropertyValue(m_Target, L"LayerThickness", {MakeIntArg(layerIndex), MakeDoubleArg(thick)} );
    }; 

    /**
     * @brief Gets the number of elements in a layer.
     * @param layerIndex Index of the layer (default 1).
     * @return Number of elements.
     */
    int getNumberOfElements(int layerIndex=1){
        return variantTo<int>(GetPropertyValue(m_Target, L"NumberOfElements", {MakeIntArg(layerIndex)} ));
    }; 

    /**
     * @brief Sets the number of elements in a layer.
     * @param layerIndex Index of the layer (default 1).
     * @param numEl Number of elements to set (default 1).
     */
    void setNumberOfElements(int layerIndex=1, int numEl=1){
        SetPropertyValue(m_Target, L"NumberOfElements", {MakeIntArg(layerIndex), MakeIntArg(numEl)} );
    }; 

    /**
     * @brief Gets the name of an element in a layer.
     * @param layerIndex Index of the layer (default 1).
     * @param elementIndex Index of the element (default 1).
     * @return Element name.
     */
    std::wstring getElementName(int layerIndex=1, int elementIndex=1 ) {
        return variantTo<std::wstring>(GetPropertyValue(m_Target, L"ElementName", {MakeIntArg(layerIndex), MakeIntArg(elementIndex)} ));
    }; 

    /**
     * @brief Sets the name of an element in a layer.
     * @param layerIndex Index of the layer.
     * @param elementIndex Index of the element.
     * @param elname Element name (default empty).
     */
    void setElementName(int layerIndex=1, int elementIndex=1, std::wstring elname=L"" ) {
        SetPropertyValue(m_Target, L"ElementName", {MakeIntArg(layerIndex), MakeIntArg(elementIndex), MakeStringArg(elname)} );
    }; 

    /**
     * @brief Gets the concentration of an element in a layer.
     * @param layerIndex Index of the layer (default 1).
     * @param elementIndex Index of the element (default 1).
     * @return Element concentration.
     */
    double getElementConcentration(int layerIndex=1, int elementIndex=1 ) {
        return variantTo<double>(GetPropertyValue(m_Target, L"ElementConcentration", {MakeIntArg(layerIndex), MakeIntArg(elementIndex)} ));
    }; 

    /**
     * @brief Sets the concentration of an element in a layer.
     * @param layerIndex Index of the layer.
     * @param elementIndex Index of the element.
     * @param conc Concentration to set (default 0.0).
     */
    void setElementConcentration(int layerIndex=1, int elementIndex=1, double conc=0.0) {
        SetPropertyValue(m_Target, L"ElementConcentration", {MakeIntArg(layerIndex), MakeIntArg(elementIndex), MakeDoubleArg(conc)} );
    }; 

    /// @}

    /// @name SIMNRA.Setup methods
    /// @{

    /**
     * @brief Gets the full width at half maximum (FWHM) of the beam energy.
     * @return FWHM value.
     */
    double getBeamEnergyFWHM( ) {
        return variantTo<double>(GetPropertyValue(m_Setup, L"Beamspread", {} ));
    }; 

    /**
     * @brief Sets the FWHM of the beam energy.
     * @param fwhm The value to set.
     */
    void setBeamEnergyFWHM(double fwhm) {
        SetPropertyValue(m_Setup, L"Beamspread", {MakeDoubleArg(fwhm)} );
    }; 

    /**
     * @brief Gets the beam energy.
     * @return Beam energy value.
     */
    double getBeamEnergy( ) {
        return variantTo<double>(GetPropertyValue(m_Setup, L"Energy", {} ));
    }; 

    /**
     * @brief Sets the beam energy.
     * @param E Energy value to set.
     */
    void setBeamEnergy(double E) {
        SetPropertyValue(m_Setup, L"Energy", {MakeDoubleArg(E)} );
    }; 
    /// @}


};
    

#endif
