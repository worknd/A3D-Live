//===========================================================================
//
// A3D_DAL.CPP
//
// Purpose: Emulate Aureal A3D DAL interface (Wrapper to DirectSound3D).
//
// Copyright (C) 2004 Dmitry Nesterenko. All rights reserved.
//
//===========================================================================


#include <math.h>
#include <dsound.h>


#include "ia3dapi.h"
#include "a3d_dll.h"
#include "a3d_dal.h"
#include "a3d_ref.h"


#ifdef _DEBUG
#	include <crtdbg.h>
#else
#	pragma intrinsic(memset, memcmp, memcpy, strcpy, strcat, strlen, sin, cos, log10)
#endif


// Compilation options.
#ifdef _DEBUG
#	define EXTENDED_LOG TRUE
#endif


//===========================================================================
//
// IA3dScaleHack::IA3dScaleHack
// IA3dScaleHack::~IA3dScaleHack
//
// Constructor Parameters:
//  None
//
//===========================================================================
IA3dScaleHack::IA3dScaleHack() :
	m_cRef(0),
	m_bScaleHack(A3D_FALSE)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dScaleHack::IA3dScaleHack()=%u"), g_cObj + 1);
#endif
	// Increase object counter.
	InterlockedIncrement(&g_cObj);
}

IA3dScaleHack::~IA3dScaleHack()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dScaleHack::~IA3dScaleHack()=%u"), g_cObj - 1);
	_ASSERTE(g_cObj > 0);
#endif
	// Decrease object counter.
	InterlockedDecrement(&g_cObj);
}


//===========================================================================
//
// IA3dScaleHack::QueryInterface
// IA3dScaleHack::AddRef
// IA3dScaleHack::Release
//
// Purpose: Standard OLE routines needed for all interfaces.
//
//===========================================================================
STDMETHODIMP IA3dScaleHack::QueryInterface(REFIID rIid, LPVOID * ppvObj)
{
#ifdef _DEBUG
	TCHAR szGuid[64];
	LogMsg(TEXT("IA3dScaleHack::QueryInterface(%s,%#x)"),
		GuidToStr(rIid, szGuid, sizeof(szGuid)), ppvObj);
	_ASSERTE(ppvObj && !IsBadWritePtr(ppvObj, sizeof(*ppvObj)));
#endif
	// Check arguments values.
	if (!ppvObj)
		return E_POINTER;

	// For future invalid return.
	*ppvObj = NULL;

#ifdef _DEBUG
	if ((IID_IUnknown != rIid) && (IID_IA3dScaleHack != rIid))
		LogMsg(TEXT("...QueryInterface()=E_NOINTERFACE!"));
#endif
	// Only Unknown or A3dScaleHack objects.
	if (IID_IUnknown != rIid && IID_IA3dScaleHack != rIid)
		return E_NOINTERFACE;

	// Return this object;
	*ppvObj = this;
	AddRef();
	
	return S_OK;
}

STDMETHODIMP_(ULONG) IA3dScaleHack::AddRef()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dScaleHack::AddRef()=%u"), m_cRef + 1);
	_ASSERTE(m_cRef >= 0);
#endif
	// Increase reference counter.
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) IA3dScaleHack::Release()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dScaleHack::Release()=%u"), m_cRef - 1);
	_ASSERTE(m_cRef > 0);
#endif
	// Decrease reference counter.
	ULONG cRef = InterlockedDecrement(&m_cRef);
	if (!cRef)
	{
		delete this;
		return 0;
	}

	return cRef;
}


//===========================================================================
//
// IA3dScaleHack::ScaleHackEnable
//
// Purpose: Enable A3D scale hack.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dScaleHack::ScaleHackEnable()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dScaleHack::ScaleHackEnable()"));
#endif
	// Set A3D scale hack value.
	m_bScaleHack = A3D_TRUE;

	return S_OK;
}

//===========================================================================
//
// IA3dScaleHack::ScaleHackDisable
//
// Purpose: Disable A3D scale hack.
//
// Return: S_OK always.
//
//===========================================================================
STDMETHODIMP IA3dScaleHack::ScaleHackDisable()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dScaleHack::ScaleHackDisable()"));
#endif
	// Clear A3D scale hack value.
	m_bScaleHack = A3D_FALSE;

	return S_OK;
}

//===========================================================================
//
// IA3dScaleHack::IsScaleHackNeeded
//
// Purpose: Check need A3D scale hack.
//
// Parameters:
//  pbScaleIsNeeded LPBOOL pointer to buffer for flag need A3D scale hack.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dScaleHack::IsScaleHackNeeded(LPBOOL pbScaleIsNeeded)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dScaleHack::IsScaleHackNeeded(%#x)"), pbScaleIsNeeded);
	_ASSERTE(pbScaleIsNeeded && !IsBadWritePtr(pbScaleIsNeeded, sizeof(*pbScaleIsNeeded)));
#endif
	// Check arguments values.
	if (!pbScaleIsNeeded)
		return E_POINTER;

	// Scale hack needed always.
	*pbScaleIsNeeded = A3D_TRUE;

	return S_OK;
}


//===========================================================================
//
// IA3dScaleHackBuffer::IA3dScaleHackBuffer
// IA3dScaleHackBuffer::~IA3dScaleHackBuffer
//
// Constructor Parameters:
//  None
//
//===========================================================================
IA3dScaleHackBuffer::IA3dScaleHackBuffer() :
	m_cRef(0),
	m_bScaleHackBuffer(A3D_FALSE)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dScaleHackBuffer::IA3dScaleHackBuffer()=%u"), g_cObj + 1);
#endif
	// Increase object counter.
	InterlockedIncrement(&g_cObj);
}

IA3dScaleHackBuffer::~IA3dScaleHackBuffer()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dScaleHackBuffer::~IA3dScaleHackBuffer()=%u"), g_cObj - 1);
	_ASSERTE(g_cObj > 0);
#endif
	// Decrease object counter.
	InterlockedDecrement(&g_cObj);
}


//===========================================================================
//
// IA3dScaleHackBuffer::QueryInterface
// IA3dScaleHackBuffer::AddRef
// IA3dScaleHackBuffer::Release
//
// Purpose: Standard OLE routines needed for all interfaces.
//
//===========================================================================
STDMETHODIMP IA3dScaleHackBuffer::QueryInterface(REFIID rIid, LPVOID * ppvObj)
{
#ifdef _DEBUG
	TCHAR szGuid[64];
	LogMsg(TEXT("IA3dScaleHackBuffer::QueryInterface(%s,%#x)"),
		GuidToStr(rIid, szGuid, sizeof(szGuid)), ppvObj);
	_ASSERTE(ppvObj && !IsBadWritePtr(ppvObj, sizeof(*ppvObj)));
#endif
	// Check arguments values.
	if (!ppvObj)
		return E_POINTER;

	// For future invalid return.
	*ppvObj = NULL;

#ifdef _DEBUG
	if ((IID_IUnknown != rIid) && (IID_IA3dScaleHackBuffer != rIid))
		LogMsg(TEXT("...QueryInterface()=E_NOINTERFACE!"));
#endif
	// Only Unknown or A3dScaleHackBuffer objects.
	if ((IID_IUnknown != rIid) && (IID_IA3dScaleHackBuffer != rIid))
		return E_NOINTERFACE;

	// Return this object;
	*ppvObj = this;
	AddRef();
	
	return S_OK;
}

STDMETHODIMP_(ULONG) IA3dScaleHackBuffer::AddRef()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dScaleHackBuffer::AddRef()=%u"), m_cRef + 1);
	_ASSERTE(m_cRef >= 0);
#endif
	// Increase reference counter.
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) IA3dScaleHackBuffer::Release()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dScaleHackBuffer::Release()=%u"), m_cRef - 1);
	_ASSERTE(m_cRef > 0);
#endif
	// Decrease reference counter.
	ULONG cRef = InterlockedDecrement(&m_cRef);
	if (!cRef)
	{
		delete this;
		return 0;
	}

	return cRef;
}


//===========================================================================
//
// IA3dScaleHackBuffer::SetScaleHack
//
// Purpose: Set flag value A3D scale hack for buffer.
//
// Parameters:
//  bScaleHackBuffer BOOL new flag value A3D scale hack for buffer.
//
// Return: S_OK always.
//
//===========================================================================
STDMETHODIMP IA3dScaleHackBuffer::SetScaleHack(BOOL bScaleHackBuffer)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dScaleHackBuffer::SetScaleHack(%u)"), bScaleHackBuffer);
#endif
	// Set new scale hack buffer flag.
	m_bScaleHackBuffer = bScaleHackBuffer;

	return S_OK;
}


//===========================================================================
//
// IA3dDal::SetA3dCaps
//
// Purpose: Set A3D DAL capabilities.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dDal::SetA3dCaps()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDal::SetA3dCaps()"));
	_ASSERTE(m_pDS);
#endif
	DSCAPS DSCaps;
	DSCaps.dwSize = sizeof(DSCaps);

	// Get DirectSound capabilities.
	HRESULT hr = m_pDS->GetCaps(&DSCaps);
	if (FAILED(hr))
		return hr;

	// Check audio driver compatibility.
	if ((DSCaps.dwFlags & DSCAPS_EMULDRIVER) ||
	!(DSCaps.dwFlags & DSCAPS_PRIMARY16BIT) ||
	!(DSCaps.dwFlags & DSCAPS_PRIMARYSTEREO))
		return E_FAIL;

	// Set main capabilities information.
	ZeroMemory(&m_A3dCaps, sizeof(m_A3dCaps));
	m_A3dCaps.wVersion = A3DCAPS_VERSION;
	m_A3dCaps.dwCalcFactor = A3DCAPS_CALC_FACTOR;
	m_A3dCaps.wChannels = 2;
	m_A3dCaps.dwMaxSampleRate = A3D_SAMPLE_RATE_1;
	m_A3dCaps.wMaxBuffers = (WORD)DSCaps.dwMaxHw3DAllBuffers;
	m_A3dCaps.dwSamplesMask = 1;

	// Enable first reflections.
	if (DSCaps.dwMaxHw3DAllBuffers && (DSCaps.dwFlags & DSCAPS_CONTINUOUSRATE))
	{
		// Set device depended information.
		m_A3dCaps.wDeviceType = A3DCAPS_DEVICE_TYPE_AU8830;
		m_A3dCaps.wValue_8 = 4;
		m_A3dCaps.wRefMask = 0x7;
		m_A3dCaps.dwValue_1Ch = A3DCAPS_1CH_AU8830;
		m_A3dCaps.dwFeatureFlags = A3DCAPS_AU8830_DEFAULT;
		m_A3dCaps.dwValue_44h = A3DCAPS_44H_AU8830;

		// Set first reflections information.
		m_A3dCaps.dwMaxReflections = DSCaps.dwMaxHw3DAllBuffers;
		m_A3dCaps.dwMaxSrcReflections = A3D_MAX_SOURCE_REFLECTIONS;
	}
	else
	{
		// Set device depended information.
		m_A3dCaps.wDeviceType = A3DCAPS_DEVICE_TYPE_AU8820;
		m_A3dCaps.wValue_8 = 2;
		m_A3dCaps.dwValue_1Ch = A3DCAPS_1CH_AU8820;
		m_A3dCaps.dwFeatureFlags = A3DCAPS_AU8820_DEFAULT;
		m_A3dCaps.dwValue_44h = A3DCAPS_44H_AU8820;
	}

	// Set first sample capabilities.
	m_A3dCaps.Samples[0].wEnable = A3D_TRUE;
	m_A3dCaps.Samples[0].dwSamplesPerSec1 = A3D_SAMPLE_RATE_1;
	m_A3dCaps.Samples[0].dwSamplesPerSec2 = A3D_SAMPLE_RATE_1;
	m_A3dCaps.Samples[0].wCtrlSize = sizeof(A3DCTRL_SRC_SUPER);
	m_A3dCaps.Samples[0].wCtrlOffsetLeft = 28;
	m_A3dCaps.Samples[0].wCtrlOffsetRight = 44;
	m_A3dCaps.Samples[0].wHrtfMode = A3DCAPS_HAWK44;
	m_A3dCaps.Samples[0].wBitsPerSample = 16;
	m_A3dCaps.Samples[0].wValue_1Ah = A3DCAPS_SAMPLE_1AH;
	m_A3dCaps.Samples[0].wValue_1Ch = A3DCAPS_SAMPLE_1CH;

	// Enable second sample.
	if (A3D_SAMPLE_RATE_2 >= DSCaps.dwMinSecondarySampleRate &&
	A3D_SAMPLE_RATE_2 <= DSCaps.dwMaxSecondarySampleRate)
	{
		// Set new base sample rate.
		m_A3dCaps.dwMaxSampleRate = A3D_SAMPLE_RATE_2;

		// Set second bit in sample mask.
		m_A3dCaps.dwSamplesMask |= 2;

		// Set second sample capabilities.
		m_A3dCaps.Samples[1].wEnable = A3D_TRUE;
		m_A3dCaps.Samples[1].dwSamplesPerSec1 = A3D_SAMPLE_RATE_2;
		m_A3dCaps.Samples[1].dwSamplesPerSec2 = A3D_SAMPLE_RATE_2;
		m_A3dCaps.Samples[1].wCtrlSize = sizeof(A3DCTRL_SRC_SUPER);
		m_A3dCaps.Samples[1].wCtrlOffsetLeft = 28;
		m_A3dCaps.Samples[1].wCtrlOffsetRight = 44;
		m_A3dCaps.Samples[1].wHrtfMode = A3DCAPS_HAWK44;
		m_A3dCaps.Samples[1].wBitsPerSample = 16;
		m_A3dCaps.Samples[1].wValue_1Ah = A3DCAPS_SAMPLE_1AH;
		m_A3dCaps.Samples[1].wValue_1Ch = A3DCAPS_SAMPLE_1CH;
	}

	return S_OK;
}


//===========================================================================
//
// IA3dDal::IA3dDal
// IA3dDal::~IA3dDal
//
// Constructor Parameters:
//  pDS             LPDIRECTSOUND to the parent object.
//
//===========================================================================
IA3dDal::IA3dDal(LPDIRECTSOUND pDS) :
	m_cRef(0),
	m_pDS(pDS)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDal::IA3dDal()=%u"), g_cObj + 1);
#endif
	// Zero big object members.
	ZeroMemory(&m_A3dCaps, sizeof(m_A3dCaps));

	// Add reference for parent DirectSound object.
	if (m_pDS)
		m_pDS->AddRef();
	else
		CoCreateInstance(CLSID_DirectSound, NULL, CLSCTX_INPROC_SERVER,
			IID_IDirectSound, (LPVOID *)&m_pDS);

	// Increase object counter.
	InterlockedIncrement(&g_cObj);
}

IA3dDal::~IA3dDal()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDal::~IA3dDal()=%u"), g_cObj - 1);
	_ASSERTE(g_cObj > 0);
#endif
	// Release parent DirectSound object.
	if (m_pDS)
		m_pDS->Release();

	// Decrease object counter.
	InterlockedDecrement(&g_cObj);
}


//===========================================================================
//
// IA3dDal::QueryInterface
// IA3dDal::AddRef
// IA3dDal::Release
//
// Purpose: Standard OLE routines needed for all interfaces.
//
//===========================================================================
STDMETHODIMP IA3dDal::QueryInterface(REFIID rIid, LPVOID * ppvObj)
{
#ifdef _DEBUG
	TCHAR szGuid[64];
	LogMsg(TEXT("IA3dDal::QueryInterface(%s,%#x)"),
		GuidToStr(rIid, szGuid, sizeof(szGuid)), ppvObj);
	_ASSERTE(ppvObj && !IsBadWritePtr(ppvObj, sizeof(*ppvObj)));
	_ASSERTE(m_pDS);
#endif
	// Check arguments values.
	if (!ppvObj)
		return E_POINTER;

	// For future invalid return.
	*ppvObj = NULL;

	// If not exist parent DirectSound object.
	if (!m_pDS)
		return E_FAIL;

	// Request for A3dScaleHack object.
	if (IID_IA3dScaleHack == rIid)
	{
		// Create A3dScaleHack object.
		LPA3DSCALEHACK pA3dScaleHack = new IA3dScaleHack;
		if (!pA3dScaleHack)
			return E_OUTOFMEMORY;

		// Kill the object if initial creation failed.
		HRESULT hr = pA3dScaleHack->QueryInterface(rIid, ppvObj);
		if (FAILED(hr))
			delete pA3dScaleHack;

		return hr;
	}

	// Request for A3d or A3d2 object.
	if (IID_IA3d == rIid || IID_IA3d2 == rIid)
	{
		// Create A3dX object.
		LPA3DX pA3dX = new IA3dX(m_pDS);
		if (!pA3dX)
			return E_OUTOFMEMORY;

		// Kill the object if initial creation or Init failed.
		HRESULT hr = pA3dX->QueryInterface(rIid, ppvObj);
		if (FAILED(hr))
			delete pA3dX;

		return hr;
	}

	// Request for DirectSound object.
	if (IID_IDirectSound == rIid)
	{
		// Create A3dSound object.
		LPA3DSOUND pA3dSound = new IA3dSound(m_pDS);
		if (!pA3dSound)
			return E_OUTOFMEMORY;

		// Kill the object if initial creation failed.
		HRESULT hr = pA3dSound->QueryInterface(rIid, ppvObj);
		if (FAILED(hr))
			delete pA3dSound;

		return hr;
	}

#ifdef _DEBUG
	if ((IID_IUnknown != rIid) && (IID_IA3dDal != rIid))
		LogMsg(TEXT("...QueryInterface()=E_NOINTERFACE!"));
#endif
	// Only Unknown or A3dDal objects.
	if ((IID_IUnknown != rIid) && (IID_IA3dDal != rIid))
		return E_NOINTERFACE;

	// Return this object;
	*ppvObj = this;
	AddRef();
	
	return S_OK;
}

STDMETHODIMP_(ULONG) IA3dDal::AddRef()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDal::AddRef()=%u"), m_cRef + 1);
	_ASSERTE(m_cRef >= 0);
#endif
	// Increase reference counter.
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) IA3dDal::Release()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDal::Release()=%u"), m_cRef - 1);
	_ASSERTE(m_cRef > 0);
#endif
	// Decrease reference counter.
	ULONG cRef = InterlockedDecrement(&m_cRef);
	if (!cRef)
	{
		delete this;
		return 0;
	}

	return cRef;
}


//===========================================================================
//
// IA3dDal::InitializeEx
//
// Purpose: Initialize A3D DAL interface.
//
// Parameters:
//  pGuidDevice     LPGUID identifying the initialized device.
//  dwFeaturesRequested DWORD requested initalization features.
//  dwFlags         DWORD initalization flags.
//  pdwFeaturesEnabled LPDWORD pointer to buffer for enabled features.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dDal::InitializeEx(LPGUID pGuidDevice, DWORD dwFeaturesRequested,
	DWORD dwFlags, LPDWORD pdwFeaturesEnabled)
{
#ifdef _DEBUG
	TCHAR szGuid[64];
	LogMsg(TEXT("IA3dDal::InitializeEx(%s,%#x,%#x,%#x)"),
		GuidToStr((REFGUID)*pGuidDevice, szGuid, sizeof(szGuid)),
		dwFeaturesRequested, dwFlags, pdwFeaturesEnabled);
	_ASSERTE(pdwFeaturesEnabled && !IsBadWritePtr(pdwFeaturesEnabled, sizeof(*pdwFeaturesEnabled)));
	_ASSERTE(m_pDS);
#endif
	// Check arguments values.
	if (!pdwFeaturesEnabled)
		return E_POINTER;

	// For future invalid return.
	*pdwFeaturesEnabled = 0;

	// Initialize parent DirectSound object.
	HRESULT hr = m_pDS->Initialize(pGuidDevice);
	if (FAILED(hr))
		return hr;

	// Get current foreground window.
	HWND hWnd = GetForegroundWindow();
	if (!hWnd)
		return E_FAIL;
	
	// Set cooperative level for parent DirectSound object.
	hr = m_pDS->SetCooperativeLevel(hWnd, DSSCL_NORMAL);
	if (FAILED(hr))
		return hr;

	// Get A3D DAL capabilities.
	hr = SetA3dCaps();
	if (FAILED(hr))
		return hr;

	// Copy enabled by default A3D features.
	*pdwFeaturesEnabled = dwFeaturesRequested & (A3D_DIRECT_PATH_A3D |
		A3D_DIRECT_PATH_GENERIC | A3D_OCCLUSIONS |
		A3D_DISABLE_SPLASHSCREEN | A3D_REVERB | A3D_DISABLE_FOCUS_MUTE);

	// Enable 1st reflections.
	if (m_A3dCaps.dwMaxReflections)
		*pdwFeaturesEnabled |= dwFeaturesRequested &
			(A3D_1ST_REFLECTIONS | A3D_GEOMETRIC_REVERB);

#ifdef _DEBUG
	LogMsg(TEXT("...Flags=%#x"), dwFlags);
	LogMsg(TEXT("...FeaturesRequested=%#x"), dwFeaturesRequested);
	LogMsg(TEXT("...FeaturesEnabled=%#x"), *pdwFeaturesEnabled);
#endif
	return S_OK;
}


//===========================================================================
//
// IA3dDal::CreateSoundBufferEx
//
// Purpose: Create new DirectSoundBuffer object.
//
// Parameters:
//  pcDSBufferDesc  LPCDSBUFFERDESC notification events count.
//  pbWave          LPBYTE pointer to wave data area.
//  ppDirectSoundBuffer LPDIRECTSOUNDBUFFER* pointer to buffer for DirectSoundBuffer.
//  pUnkOuter       LPUNKNOWN pointer to the controlling IUnknown.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dDal::CreateSoundBufferEx(LPCDSBUFFERDESC pcDSBufferDesc,
	LPBYTE pbWave, LPDIRECTSOUNDBUFFER * ppDirectSoundBuffer, LPUNKNOWN pUnkOuter)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDal::CreateSoundBufferEx(%#x,%#x,%#x,%#x)"),
		pcDSBufferDesc, pbWave, ppDirectSoundBuffer, pUnkOuter);
	_ASSERTE(pcDSBufferDesc && !IsBadReadPtr(pcDSBufferDesc, sizeof(*pcDSBufferDesc)));
	LogMsg(TEXT("...Size=%u"), pcDSBufferDesc->dwSize);
	LogMsg(TEXT("...Flags=%#x"), pcDSBufferDesc->dwFlags);
	if (pcDSBufferDesc->dwFlags & DSBCAPS_PRIMARYBUFFER)
		LogMsg(TEXT("......DSBCAPS_PRIMARYBUFFER"));
	if (pcDSBufferDesc->dwFlags & DSBCAPS_CTRL3D)
		LogMsg(TEXT("......DSBCAPS_CTRL3D"));
	if (pcDSBufferDesc->dwFlags & DSBCAPS_LOCHARDWARE)
		LogMsg(TEXT("......DSBCAPS_LOCHARDWARE"));
	if (pcDSBufferDesc->dwFlags & DSBCAPS_LOCSOFTWARE)
		LogMsg(TEXT("......DSBCAPS_LOCSOFTWARE"));
	LogMsg(TEXT("...BufferBytes=%u"), pcDSBufferDesc->dwBufferBytes);
	if (pcDSBufferDesc->dwReserved)
		LogMsg(TEXT("...Reserved=%#x"), pcDSBufferDesc->dwReserved);
	if (pcDSBufferDesc->lpwfxFormat)
	{
		LogMsg(TEXT("...Channels=%u"), pcDSBufferDesc->lpwfxFormat->nChannels);
		LogMsg(TEXT("...SamplesPerSec=%u"), pcDSBufferDesc->lpwfxFormat->nSamplesPerSec);
		LogMsg(TEXT("...AvgBytesPerSec=%u"), pcDSBufferDesc->lpwfxFormat->nAvgBytesPerSec);
		LogMsg(TEXT("...BlockAlign=%u"), pcDSBufferDesc->lpwfxFormat->nBlockAlign);
		LogMsg(TEXT("...BitsPerSample=%u"), pcDSBufferDesc->lpwfxFormat->wBitsPerSample);
	}
	_ASSERTE(pbWave && !IsBadReadPtr(pbWave, sizeof(*pbWave)));
	_ASSERTE(ppDirectSoundBuffer && !IsBadWritePtr(ppDirectSoundBuffer, sizeof(*ppDirectSoundBuffer)));
	_ASSERTE(!pUnkOuter);
	_ASSERTE(m_pDS);
#endif
	// Check arguments values.
	if (!pcDSBufferDesc || !pbWave || !ppDirectSoundBuffer)
		return E_POINTER;

	// For future invalid return.
	*ppDirectSoundBuffer = NULL;

	// This object doesnt support aggregation.
	if (pUnkOuter)
		return CLASS_E_NOAGGREGATION;

#if A3D_EMULATE_MAXIMUM
	DWORD dwUseDSB;

	// Get functionality A3D DAL interface.
	HRESULT hr = QueryFunctionality(A3DFUNC_USE_DS_BUFFERS, &dwUseDSB);
	if (FAILED(hr))
		return hr;

	// Create new DirectSoundBuffer object.
	if (dwUseDSB && pcDSBufferDesc->lpwfxFormat &&
	!(pcDSBufferDesc->dwFlags & DSBCAPS_PRIMARYBUFFER) &&
	pcDSBufferDesc->lpwfxFormat->nChannels > 1 &&
	(pcDSBufferDesc->dwFlags & DSBCAPS_CTRL3D))
	{
		// Use as DirectSound software buffer.
		DSBUFFERDESC DSBufDesc;
		CopyMemory(&DSBufDesc, pcDSBufferDesc, sizeof(DSBufDesc));
		DSBufDesc.dwFlags = (DSBufDesc.dwFlags & ~DSBCAPS_LOCHARDWARE) | DSBCAPS_LOCSOFTWARE;
		return m_pDS->CreateSoundBuffer(&DSBufDesc, ppDirectSoundBuffer, NULL);
	}

	// Use as standart DirectSound buffer.
	if (!(pcDSBufferDesc->dwFlags & DSBCAPS_CTRL3D) || 
	(pcDSBufferDesc->dwFlags & DSBCAPS_LOCSOFTWARE))
		return m_pDS->CreateSoundBuffer(pcDSBufferDesc, ppDirectSoundBuffer, NULL);

	// Use hardware sound buffers, only if they available.
	DSBUFFERDESC DSBufDesc;
	CopyMemory(&DSBufDesc, pcDSBufferDesc, sizeof(DSBufDesc));
	DSBufDesc.dwSize = sizeof(DSBufDesc);
	DSBufDesc.dwFlags |= DSBCAPS_CTRLVOLUME | DSBCAPS_CTRLFREQUENCY |
		DSBCAPS_CTRLPOSITIONNOTIFY;
#ifdef _DEBUG
		LogMsg(TEXT("...Flags=%#x!"), DSBufDesc.dwFlags);
#endif
	DSBufDesc.guid3DAlgorithm = DS3DALG_HRTF_FULL;
	hr = m_pDS->CreateSoundBuffer(&DSBufDesc, ppDirectSoundBuffer, NULL);
	if (SUCCEEDED(hr))
	{
		// Create new A3dSoundBuffer object.
		LPA3DSOUNDBUFFER pA3dSoundBuffer = new IA3dSoundBuffer(*ppDirectSoundBuffer, m_pDS);
		if (!pA3dSoundBuffer)
		{
			(*ppDirectSoundBuffer)->Release();
			*ppDirectSoundBuffer = NULL;
			return E_OUTOFMEMORY;
		}

		// Kill the object if initial creation failed.
		hr = pA3dSoundBuffer->QueryInterface(IID_IDirectSoundBuffer, (LPVOID *)ppDirectSoundBuffer);
		if (FAILED(hr))
			delete pA3dSoundBuffer;
	}

#ifdef _DEBUG
	LogMsg(TEXT("...=%s"), Result(hr));
#endif
	return hr;
#else
	// Not implemented.
	return E_NOTIMPL;
#endif
}


//===========================================================================
//
// IA3dDal::GetA3dCaps
//
// Purpose: Get A3D DAL capabilities.
//
// Parameters:
//  pA3dCaps        LPA3D_CAPS pointer to buffer for get capabilities.
//  pdwSize         LPDWORD pointer to buffer for capabilities buffer size.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dDal::GetA3dCaps(LPA3D_CAPS pA3dCaps, LPDWORD pdwSize)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDal::GetA3dCaps(%#x,%#x)"), pA3dCaps, pdwSize);
	_ASSERTE(pA3dCaps && !IsBadWritePtr(pA3dCaps, sizeof(*pA3dCaps)));
	_ASSERTE(pdwSize && !IsBadWritePtr(pdwSize, sizeof(*pdwSize)));
	_ASSERTE(m_pDS);
#endif
	// Check arguments values.
	if (!pA3dCaps || !pdwSize)
		return E_POINTER;
#ifdef _DEBUG
	LogMsg(TEXT("...sizeof(A3D_CAPS)=%u(%u)"), sizeof(A3D_CAPS), *pdwSize);
#endif

	// Get A3D DAL capabilities.
	if (!m_A3dCaps.wDeviceType)
	{
		HRESULT hr = SetA3dCaps();
		if (FAILED(hr))
			return hr;
	}

	// Copy A3D DAL capabilities.
	CopyMemory(pA3dCaps, &m_A3dCaps, min(*pdwSize, sizeof(m_A3dCaps)));
	*pdwSize = sizeof(m_A3dCaps);
#if defined(_DEBUG) && defined(EXTENDED_LOG)
#define A3D_CAPS_OFF(e) ((UINT)&pA3dCaps->e - (UINT)pA3dCaps)
#define A3D_SAMPLE_OFF(i,e) ((UINT)&pA3dCaps->Samples[i].e - (UINT)&pA3dCaps->Samples[i])
	LogMsg(TEXT("...wDeviceType(%x)=%#x"), A3D_CAPS_OFF(wDeviceType), pA3dCaps->wDeviceType);
	LogMsg(TEXT("...wVersion(%x)=%#x"), A3D_CAPS_OFF(wVersion), pA3dCaps->wVersion);
	LogMsg(TEXT("...dwCalcFactor(%x)=%u"), A3D_CAPS_OFF(dwCalcFactor), pA3dCaps->dwCalcFactor);
	LogMsg(TEXT("...wValue_%x=%u"), A3D_CAPS_OFF(wValue_8), pA3dCaps->wValue_8);
	LogMsg(TEXT("...wChannels(%x)=%u"), A3D_CAPS_OFF(wChannels), pA3dCaps->wChannels);
	LogMsg(TEXT("...wRefMask(%x)=%#x"), A3D_CAPS_OFF(wRefMask), pA3dCaps->wRefMask);
	LogMsg(TEXT("...dwMaxSampleRate(%x)=%u"), A3D_CAPS_OFF(dwMaxSampleRate), pA3dCaps->dwMaxSampleRate);
	LogMsg(TEXT("...wMaxBuffers(%x)=%u"), A3D_CAPS_OFF(wMaxBuffers), pA3dCaps->wMaxBuffers);
	LogMsg(TEXT("...dwSamplesMask(%x)=%#x"), A3D_CAPS_OFF(dwSamplesMask), pA3dCaps->dwSamplesMask);
	LogMsg(TEXT("...dwValue_%Xh=%#x"), A3D_CAPS_OFF(dwValue_1Ch), pA3dCaps->dwValue_1Ch);
	LogMsg(TEXT("...dwFeatureFlags(%x)=%#x"), A3D_CAPS_OFF(dwFeatureFlags), pA3dCaps->dwFeatureFlags);
	LogMsg(TEXT("...dwMaxReflections(%x)=%u"), A3D_CAPS_OFF(dwMaxReflections), pA3dCaps->dwMaxReflections);
	LogMsg(TEXT("...dwMaxSrcReflections(%x)=%u"), A3D_CAPS_OFF(dwMaxSrcReflections), pA3dCaps->dwMaxSrcReflections);
	LogMsg(TEXT("...dwValue_%Xh=%#x"), A3D_CAPS_OFF(dwValue_30h), pA3dCaps->dwValue_30h);
	LogMsg(TEXT("...dwValue_%Xh=%#x"), A3D_CAPS_OFF(dwValue_34h), pA3dCaps->dwValue_34h);
	LogMsg(TEXT("...dwValue_%Xh=%#x"), A3D_CAPS_OFF(dwValue_38h), pA3dCaps->dwValue_38h);
	LogMsg(TEXT("...dwValue_%Xh=%#x"), A3D_CAPS_OFF(dwValue_44h), pA3dCaps->dwValue_44h);
	LogMsg(TEXT("...sizeof(A3D_SAMPLE_CAPS)=%u"), sizeof(A3D_SAMPLE_CAPS));

	for (UINT i = 0; i < A3DCAPS_MAX_SAMPLES; i++)
	{
		if (pA3dCaps->Samples[i].wEnable)
		{
			LogMsg(TEXT("...Samples[%u]=%#x(%u)"), i, A3D_CAPS_OFF(Samples[i]),
				sizeof(pA3dCaps->Samples[i]));
			LogMsg(TEXT("...Samples[%u].wEnable(%x)=%u"), i,
				A3D_SAMPLE_OFF(i,wEnable), pA3dCaps->Samples[i].wEnable);
			LogMsg(TEXT("...Samples[%u].dwSamplesPerSec1(%x)=%u"), i,
				A3D_SAMPLE_OFF(i,dwSamplesPerSec1), pA3dCaps->Samples[i].dwSamplesPerSec1);
			LogMsg(TEXT("...Samples[%u].dwSamplesPerSec2(%x)=%u"), i,
				A3D_SAMPLE_OFF(i,dwSamplesPerSec2), pA3dCaps->Samples[i].dwSamplesPerSec2);
			LogMsg(TEXT("...Samples[%u].wCtrlSize(%x)=%u"), i,
				A3D_SAMPLE_OFF(i,wCtrlSize), pA3dCaps->Samples[i].wCtrlSize);
			LogMsg(TEXT("...Samples[%u].wCtrlOffsetLeft(%x)=%#x"), i,
				A3D_SAMPLE_OFF(i,wCtrlOffsetLeft), pA3dCaps->Samples[i].wCtrlOffsetLeft);
			LogMsg(TEXT("...Samples[%u].wValue_%Xh=%#x"), i,
				A3D_SAMPLE_OFF(i,wValue_10h), pA3dCaps->Samples[i].wValue_10h);
			LogMsg(TEXT("...Samples[%u].wCtrlOffsetRight(%x)=%#x"), i,
				A3D_SAMPLE_OFF(i,wCtrlOffsetRight), pA3dCaps->Samples[i].wCtrlOffsetRight);
			LogMsg(TEXT("...Samples[%u].wValue_%Xh=%#x"), i,
				A3D_SAMPLE_OFF(i,wValue_14h), pA3dCaps->Samples[i].wValue_14h);
			LogMsg(TEXT("...Samples[%u].wHrtfMode(%x)=%u"), i,
				A3D_SAMPLE_OFF(i,wHrtfMode), pA3dCaps->Samples[i].wHrtfMode);
			LogMsg(TEXT("...Samples[%u].wBitsPerSample(%x)=%u"), i,
				A3D_SAMPLE_OFF(i,wBitsPerSample), pA3dCaps->Samples[i].wBitsPerSample);
			LogMsg(TEXT("...Samples[%u].wValue_%Xh=%#x"), i,
				A3D_SAMPLE_OFF(i,wValue_1Ah), pA3dCaps->Samples[i].wValue_1Ah);
			LogMsg(TEXT("...Samples[%u].wValue_%Xh=%#x"), i,
				A3D_SAMPLE_OFF(i,wValue_1Ch), pA3dCaps->Samples[i].wValue_1Ch);
			LogMsg(TEXT("...Samples[%u].wValue_%Xh=%#x"), i,
				A3D_SAMPLE_OFF(i,dwValue_20h), pA3dCaps->Samples[i].dwValue_20h);
		}
	}
#endif

	return S_OK;
}


//===========================================================================
//
// IA3dDal::GetDriverInfo
//
// Purpose: Get device driver information.
//
// Parameters:
//  phA3dVxd        LPHANDLE pointer to buffer for A3D Vxd handle.
//  ppIDsDriver     LPVOID* pointer to buffer for DsDriver interface.
//  ppIA3dDriver    LPVOID* pointer to buffer for A3D driver interface.
//  pdwCardIndex    LPDWORD pointer to buffer for sound card number.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dDal::GetDriverInfo(LPHANDLE phA3dVxd, LPVOID * ppIDsDriver,
	LPVOID * ppIA3dDriver, LPDWORD pdwCardIndex)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDal::GetDriverInfo(%#x,%#x,%#x,%#x)"),
		phA3dVxd, ppIDsDriver, ppIA3dDriver, pdwCardIndex);
	_ASSERTE(phA3dVxd && !IsBadWritePtr(phA3dVxd, sizeof(*phA3dVxd)));
	_ASSERTE(ppIDsDriver && !IsBadWritePtr(ppIDsDriver, sizeof(*ppIDsDriver)));
	_ASSERTE(ppIA3dDriver && !IsBadWritePtr(ppIA3dDriver, sizeof(*ppIA3dDriver)));
	_ASSERTE(pdwCardIndex && !IsBadWritePtr(pdwCardIndex, sizeof(*pdwCardIndex)));
#endif
	// Check arguments values.
	if (!phA3dVxd || !ppIDsDriver || !ppIA3dDriver || !pdwCardIndex)
		return E_POINTER;

	// Copy driver information.
	*phA3dVxd = INVALID_HANDLE_VALUE;
	*ppIDsDriver = NULL;
	*ppIA3dDriver = NULL;
	*pdwCardIndex = 0;

	return S_OK;
}


//===========================================================================
//
// IA3dDal::GetDS
//
// Purpose: Get DirectSound interface.
//
// Parameters:
//  ppDirectSound   LPDIRECTSOUND pointer to buffer for DirectSound interface.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dDal::GetDS(LPDIRECTSOUND * ppDirectSound)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDal::GetDS(%#x)"), ppDirectSound);
	_ASSERTE(ppDirectSound && !IsBadWritePtr(ppDirectSound, sizeof(*ppDirectSound)));
#endif
	// Check arguments values.
	if (!ppDirectSound)
		return E_POINTER;

#if A3D_EMULATE_MAXIMUM
	// Copy parent DirectSound object pointer.
	*ppDirectSound = m_pDS;
#else
	// Copy parent DirectSound object pointer.
	*ppDirectSound = NULL;
#endif

	return S_OK;
}


//===========================================================================
//
// IA3dDal::GetDSDriverDesc
//
// Purpose: Get DirectSound driver description.
//
// Parameters:
//  pDSDriverDesc   LPDSDRIVERDESC pointer to buffer for get driver description.
//  pdwSize         LPDWORD pointer to buffer for driver description buffer size.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dDal::GetDSDriverDesc(LPDSDRIVERDESC pDSDriverDesc, LPDWORD pdwSize)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDal::GetDSDriverDesc(%#x,%#x)"),
		pDSDriverDesc, pdwSize);
	_ASSERTE(pDSDriverDesc && !IsBadWritePtr(pDSDriverDesc, sizeof(*pDSDriverDesc)));
	_ASSERTE(pdwSize && !IsBadWritePtr(pdwSize, sizeof(*pdwSize)));
#endif
	// Check arguments values.
	if (!pDSDriverDesc || !pdwSize)
		return E_POINTER;

#ifdef _DEBUG
	LogMsg(TEXT("...sizeof(DSDRIVERDESC)=%u(%u)"), sizeof(DSDRIVERDESC), *pdwSize);
#endif
	// Copy DirectSound driver description.
	ZeroMemory(pDSDriverDesc, min(*pdwSize, sizeof(*pDSDriverDesc)));
#if A3D_EMULATE_MAXIMUM
	pDSDriverDesc->dwFlags = 4;
	strcpy(pDSDriverDesc->szDesc, "A3D-Live!");
	strcpy(pDSDriverDesc->szDrvname, "a3d.dll");
	pDSDriverDesc->dnDevNode = 1;
	pDSDriverDesc->wVxdId = 0x3FB8;
#endif
	*pdwSize = sizeof(*pDSDriverDesc);
#if defined(_DEBUG) && defined(EXTENDED_LOG)
#define DS_DRVDESC_OFF(e) ((UINT)&pDSDriverDesc->e - (UINT)pDSDriverDesc)
	LogMsg(TEXT("...dwFlags(%x)=%#x"), DS_DRVDESC_OFF(dwFlags), pDSDriverDesc->dwFlags);
	LogMsg(TEXT("...szDesc(%x)=%s"), DS_DRVDESC_OFF(szDesc), pDSDriverDesc->szDesc);
	LogMsg(TEXT("...szDrvname(%x)=%s"), DS_DRVDESC_OFF(szDrvname), pDSDriverDesc->szDrvname);
	LogMsg(TEXT("...dnDevNode(%x)=%#x"), DS_DRVDESC_OFF(dnDevNode), pDSDriverDesc->dnDevNode);
	LogMsg(TEXT("...wVxdId(%x)=%#x"), DS_DRVDESC_OFF(wVxdId), pDSDriverDesc->wVxdId);
	LogMsg(TEXT("...ulDeviceNum(%x)=%#x"), DS_DRVDESC_OFF(ulDeviceNum), pDSDriverDesc->ulDeviceNum);
	LogMsg(TEXT("...dwHeapType(%x)=%#x"), DS_DRVDESC_OFF(dwHeapType), pDSDriverDesc->dwHeapType);
	LogMsg(TEXT("...pvDirectDrawHeap(%x)=%#x"), DS_DRVDESC_OFF(pvDirectDrawHeap), pDSDriverDesc->pvDirectDrawHeap);
	LogMsg(TEXT("...dwMemStartAddress(%x)=%#x"), DS_DRVDESC_OFF(dwMemStartAddress), pDSDriverDesc->dwMemStartAddress);
	LogMsg(TEXT("...dwMemEndAddress(%x)=%#x"), DS_DRVDESC_OFF(dwMemEndAddress), pDSDriverDesc->dwMemEndAddress);
	LogMsg(TEXT("...dwMemAllocExtra(%x)=%#x"), DS_DRVDESC_OFF(dwMemAllocExtra), pDSDriverDesc->dwMemAllocExtra);
#endif

	return S_OK;
}


//===========================================================================
//
// IA3dDal::QueryFunctionality
//
// Purpose: Query functionality A3D DAL interface.
//
// Parameters:
//  dwFunction      DWORD number of requested functionality.
//  pdwStatus       LPDWORD pointer to buffer for get status.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dDal::QueryFunctionality(DWORD dwFunction, LPDWORD pdwStatus)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDal::QueryFunctionality(%u,%#x)"), dwFunction, pdwStatus);
	_ASSERTE(pdwStatus && !IsBadWritePtr(pdwStatus, sizeof(*pdwStatus)));
#endif
	// Check arguments values.
	if (!pdwStatus)
		return E_POINTER;

#if A3D_EMULATE_MAXIMUM
	// Copy current functionality status.
	switch(dwFunction)
	{
	case A3DFUNC_0:
		*pdwStatus = (m_A3dCaps.dwFeatureFlags & 0x1) ? A3D_FALSE : A3D_TRUE;
		break;
	case A3DFUNC_1:
		*pdwStatus = (m_A3dCaps.dwFeatureFlags & 0x2) ? A3D_FALSE : A3D_TRUE;
		break;
	case A3DFUNC_2:
		*pdwStatus = (m_A3dCaps.dwFeatureFlags & 0x40) ? A3D_FALSE : A3D_TRUE;
		break;
	case A3DFUNC_3:
		*pdwStatus = (m_A3dCaps.dwFeatureFlags & 0x800) ? A3D_TRUE : A3D_FALSE;
		break;
	case A3DFUNC_4:
		*pdwStatus = m_A3dCaps.wMaxBuffers;
		break;
	case A3DFUNC_NOT_EXCLUSIVE:
		*pdwStatus = (m_A3dCaps.dwFeatureFlags & A3DCAPS_ALL_PRIORITY) ? A3D_TRUE : A3D_FALSE;
		break;
	case A3DFUNC_6:
		*pdwStatus = (m_A3dCaps.dwFeatureFlags & 0x4000) ? A3D_TRUE : A3D_FALSE;
		break;
	case A3DFUNC_NO_SPLASHSCREEN:
		*pdwStatus = (m_A3dCaps.dwFeatureFlags & A3DCAPS_NO_SPLASHSCREEN) ? A3D_TRUE : A3D_FALSE;
		break;
	case A3DFUNC_8:
		*pdwStatus = (m_A3dCaps.dwValue_1Ch & 0x4) ? A3D_TRUE : A3D_FALSE;
		break;
	case A3DFUNC_9:
		*pdwStatus = (m_A3dCaps.dwFeatureFlags & 0x8000) ? A3D_TRUE : A3D_FALSE;
		break;
	case A3DFUNC_A:
		*pdwStatus = (m_A3dCaps.wRefMask & 0x7) ? A3D_TRUE : A3D_FALSE;
		break;
	case A3DFUNC_B:
		*pdwStatus = A3D_FALSE;
		break;
	case A3DFUNC_C:
		*pdwStatus = (m_A3dCaps.dwValue_1Ch & 0x1) ? A3D_TRUE : A3D_FALSE;
		break;
	case A3DFUNC_D:
		*pdwStatus = A3D_FALSE;
		break;
	case A3DFUNC_USE_DS_BUFFERS:
		*pdwStatus = (m_A3dCaps.dwFeatureFlags & A3DCAPS_USE_DS_BUFFERS) ? A3D_TRUE : A3D_FALSE;
		break;
	case A3DFUNC_NEED_SCALE_HACK:
		*pdwStatus = (m_A3dCaps.dwFeatureFlags & A3DCAPS_NEED_SCALE_HACK) ? A3D_TRUE : A3D_FALSE;
		break;
	default:
		*pdwStatus = A3D_FALSE;
		return E_NOTIMPL;
	}
#else
	*pdwStatus = 0;
#endif

	return S_OK;
}


//===========================================================================
//
// IA3dDal::Verify
//
// Purpose: Verify authorization for A3D DAL interface.
//
// Parameters:
//  pString         LPSTR pointer to string.
//  ppStringCrypted LPSTR* pointer to buffer for crypted string pinter.
//  ppCopyright     LPSTR* pointer to buffer for copyright string pointer.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dDal::Verify(LPSTR pString, LPSTR * ppStringCrypted,
	LPSTR * ppCopyright)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDal::Verify(%#x,%#x,%#x)"), pString, ppStringCrypted,
		ppCopyright);
	_ASSERTE(pString && !IsBadReadPtr(pString, sizeof(*pString)));
	_ASSERTE(ppStringCrypted && !IsBadReadPtr(ppStringCrypted, sizeof(*ppStringCrypted)));
	_ASSERTE(ppCopyright && !IsBadReadPtr(ppCopyright, sizeof(*ppCopyright)));
#endif
	// Check arguments values.
	if (!pString || !ppStringCrypted || !ppCopyright)
		return E_POINTER;

	// Not implemented.
	return E_NOTIMPL;
}


//===========================================================================
//
// IA3dDalBuffer::IA3dDalBuffer
// IA3dDalBuffer::~IA3dDalBuffer
//
// Constructor Parameters:
//  pDSB            LPDIRECTSOUNDBUFFER to the parent object.
//
//===========================================================================
IA3dDalBuffer::IA3dDalBuffer(LPDIRECTSOUND pDS, LPDIRECTSOUNDBUFFER pDSB) :
	m_cRef(0),
	m_pDS(pDS),
	m_pDSB(pDSB),
	m_pDS3DB(NULL),
	m_pA3dReflections(NULL)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDalBuffer::IA3dDalBuffer()=%u"), g_cObj + 1);
#endif
	// Zero big object members.
	ZeroMemory(&m_A3dCtrlSuper, sizeof(m_A3dCtrlSuper));

	// Add reference for parent DirectSound object.
	if (m_pDS)
		m_pDS->AddRef();

	// Add reference for parent DirectSoundBuffer object.
	if (m_pDSB)
		m_pDSB->AddRef();

	// Create new A3dReflections object.
	if (m_pDS && m_pDSB)
	{
		m_pA3dReflections = new IA3dReflections;
		if (m_pA3dReflections)
		{
			// Initialize A3dReflections object.
			if (FAILED(m_pA3dReflections->Initialize(m_pDS, m_pDSB)))
			{
#ifdef _DEBUG
				LogMsg(TEXT("...m_pA3dReflections->Initialize() failed!"));
#endif
				delete m_pA3dReflections;
				m_pA3dReflections = NULL;
			}
		}
	}

	// Increase object counter.
	InterlockedIncrement(&g_cObj);
}

IA3dDalBuffer::~IA3dDalBuffer()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDalBuffer::~IA3dDalBuffer()=%u"), g_cObj - 1);
	_ASSERTE(g_cObj > 0);
#endif
	// Delete A3dReflections object.
	if (m_pA3dReflections)
		delete m_pA3dReflections;

	// Release parent DirectSoundBuffer object.
	if (m_pDSB)
	{
		// Release DirectSound3DBuffer object.
		if (m_pDS3DB)
			m_pDS3DB->Release();

		m_pDSB->Release();
	}

	// Release parent DirectSound object.
	if (m_pDS)
		m_pDS->Release();

	// Decrease object counter.
	InterlockedDecrement(&g_cObj);
}


//===========================================================================
//
// IA3dDalBuffer::QueryInterface
// IA3dDalBuffer::AddRef
// IA3dDalBuffer::Release
//
// Purpose: Standard OLE routines needed for all interfaces.
//
//===========================================================================
STDMETHODIMP IA3dDalBuffer::QueryInterface(REFIID rIid, LPVOID * ppvObj)
{
#ifdef _DEBUG
	TCHAR szGuid[64];
	LogMsg(TEXT("IA3dDalBuffer::QueryInterface(%s,%#x)"),
		GuidToStr(rIid, szGuid, sizeof(szGuid)), ppvObj);
	_ASSERTE(ppvObj && !IsBadWritePtr(ppvObj, sizeof(*ppvObj)));
	_ASSERTE(m_pDSB);
#endif
	// Check arguments values.
	if (!ppvObj)
		return E_POINTER;

	// For future invalid return.
	*ppvObj = NULL;

	// If not exist parent DirectSoundBuffer object.
	if (!m_pDSB)
		return E_FAIL;

	// Request for A3dScaleHackBuffer object.
	if (IID_IA3dScaleHackBuffer == rIid)
	{
		// Create A3dScaleHackBuffer object.
		LPA3DSCALEHACKBUFFER pA3dScaleHackBuffer = new IA3dScaleHackBuffer;
		if (!pA3dScaleHackBuffer)
			return E_OUTOFMEMORY;

		// Kill the object if initial creation failed.
		HRESULT hr = pA3dScaleHackBuffer->QueryInterface(rIid, ppvObj);
		if (FAILED(hr))
			delete pA3dScaleHackBuffer;

		return hr;
	}

	// Request for DirectSoundBuffer object.
	if (IID_IDirectSoundBuffer == rIid)
	{
		// Add reference for parent DirectSoundBuffer object.
		m_pDSB->AddRef();

		// Create A3dSoundBuffer object.
		LPA3DSOUNDBUFFER pA3dSoundBuffer = new IA3dSoundBuffer(m_pDSB, m_pDS);
		if (!pA3dSoundBuffer)
		{
			m_pDSB->Release();
			return E_OUTOFMEMORY;
		}

		// Kill the object if initial creation failed.
		HRESULT hr = pA3dSoundBuffer->QueryInterface(rIid, ppvObj);
		if (FAILED(hr))
			delete pA3dSoundBuffer;

		return hr;
	}

#ifdef _DEBUG
	if ((IID_IUnknown != rIid) && (IID_IA3dDalBuffer != rIid) && (IID_IA3dDalBuffer2 != rIid))
		LogMsg(TEXT("...QueryInterface()=E_NOINTERFACE!"));
#endif
	// Only Unknown, A3dDalBuffer or A3dDalBuffer2 objects.
	if ((IID_IUnknown != rIid) && (IID_IA3dDalBuffer != rIid) && (IID_IA3dDalBuffer2 != rIid))
		return E_NOINTERFACE;

	// Return this object;
	*ppvObj = this;
	AddRef();

	return S_OK;
}

STDMETHODIMP_(ULONG) IA3dDalBuffer::AddRef()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDalBuffer::AddRef()=%u"), m_cRef + 1);
	_ASSERTE(m_cRef >= 0);
#endif
	// Increase reference counter.
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) IA3dDalBuffer::Release()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDalBuffer::Release()=%u"), m_cRef - 1);
	_ASSERTE(m_cRef > 0);
#endif
	// Decrease reference counter.
	ULONG cRef = InterlockedDecrement(&m_cRef);
	if (!cRef)
	{
		delete this;
		return 0;
	}

	return cRef;
}


//===========================================================================
//
// IA3dDalBuffer::SetA3dSuperCtrl
//
// Purpose: Set A3D DAL control packet information.
//
// Parameters:
//  pA3dCtrlSuper   LPA3DCTRL_SRC_SUPER pointer to control packet.
//  dwSize          DWORD control packet size.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dDalBuffer::SetA3dSuperCtrl(LPA3DCTRL_SRC_SUPER pA3dCtrlSuper,
	DWORD dwSize)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDalBuffer::SetA3dSuperCtrl(%#x,%u)"), pA3dCtrlSuper, dwSize);
	_ASSERTE(pA3dCtrlSuper && !IsBadReadPtr(pA3dCtrlSuper, sizeof(*pA3dCtrlSuper)));
	_ASSERTE(dwSize);
	_ASSERTE(m_pDSB);
#endif
	// Check arguments values.
	if (!pA3dCtrlSuper)
		return E_POINTER;

	// Check arguments values.
	if (!dwSize)
		return E_INVALIDARG;
#ifdef _DEBUG
	LogMsg(TEXT("...sizeof(A3DCTRL_SRC_SUPER)=%u(%u)"), sizeof(A3DCTRL_SRC_SUPER), dwSize);
#endif

	HRESULT hr;

	// Previous source control packet equal current packet.
	if (RtlEqualMemory(&m_A3dCtrlSuper, pA3dCtrlSuper, min(dwSize, sizeof(m_A3dCtrlSuper))))
	{
		// Track delay for 1st reflections.
		if (m_pA3dReflections)
		{
			hr = m_pA3dReflections->TrackDelay();
#ifdef _DEBUG
			LogMsg(TEXT("...m_pA3dReflections->TrackDelay()=%s"), Result(hr));
#endif
			if (FAILED(hr))
			{
				// Clear previous source control packet.
				ZeroMemory(&m_A3dCtrlSuper, sizeof(m_A3dCtrlSuper));
				return hr;
			}
		}

		return S_OK;
	}

#if defined(_DEBUG) && defined(EXTENDED_LOG)
#define A3D_CTRL_OFF(a) ((UINT)&pA3dCtrlSuper->a - (UINT)pA3dCtrlSuper)
#define A3D_REFS_OFF(a) ((UINT)&pA3dCtrlSuper->Reflections[i].a - (UINT)&pA3dCtrlSuper->Reflections[i])
	LogMsg(TEXT("...dwValue_%x=%#x"), A3D_CTRL_OFF(dwValue_0), pA3dCtrlSuper->dwValue_0);
	LogMsg(TEXT("...dwMode(%x)=%#x"), A3D_CTRL_OFF(dwMode), pA3dCtrlSuper->dwMode);
	LogMsg(TEXT("...bExecuted(%x)=%u"), A3D_CTRL_OFF(bExecuted), pA3dCtrlSuper->bExecuted);
	LogMsg(TEXT("...fPitch(%x)=%g"), A3D_CTRL_OFF(fPitch), pA3dCtrlSuper->fPitch);
	LogMsg(TEXT("...fAlpha(%x)=%g"), A3D_CTRL_OFF(fAlpha), pA3dCtrlSuper->fAlpha);
	LogMsg(TEXT("...fFreqFactor(%x)=%g"), A3D_CTRL_OFF(fFreqFactor), pA3dCtrlSuper->fFreqFactor);
	LogMsg(TEXT("...fNative(%x)=%g"), A3D_CTRL_OFF(fNative), pA3dCtrlSuper->fNative);
	LogMsg(TEXT("...fLeftEarAzim(%x)=%g"), A3D_CTRL_OFF(LeftEar.fAzim), pA3dCtrlSuper->LeftEar.fAzim);
	LogMsg(TEXT("...fLeftEarElev(%x)=%g"), A3D_CTRL_OFF(LeftEar.fElev), pA3dCtrlSuper->LeftEar.fElev);
	LogMsg(TEXT("...fLeftGain(%x)=%g"), A3D_CTRL_OFF(LeftEar.fGain), pA3dCtrlSuper->LeftEar.fGain);
	LogMsg(TEXT("...fLeftDelay(%x)=%g"), A3D_CTRL_OFF(LeftEar.fDelay), pA3dCtrlSuper->LeftEar.fDelay);
	LogMsg(TEXT("...fRightEarAzim(%x)=%g"), A3D_CTRL_OFF(RightEar.fAzim), pA3dCtrlSuper->RightEar.fAzim);
	LogMsg(TEXT("...fRightEarElev(%x)=%g"), A3D_CTRL_OFF(RightEar.fElev), pA3dCtrlSuper->RightEar.fElev);
	LogMsg(TEXT("...fRightGain(%x)=%g"), A3D_CTRL_OFF(RightEar.fGain), pA3dCtrlSuper->RightEar.fGain);
	LogMsg(TEXT("...fRightDelay(%x)=%g"), A3D_CTRL_OFF(RightEar.fDelay), pA3dCtrlSuper->RightEar.fDelay);
	LogMsg(TEXT("...fPriority(%x)=%g"), A3D_CTRL_OFF(fPriority), pA3dCtrlSuper->fPriority);
	LogMsg(TEXT("...fAudibility(%x)=%g"), A3D_CTRL_OFF(fAudibility), pA3dCtrlSuper->fAudibility);
	LogMsg(TEXT("...fDistance(%x)=%g"), A3D_CTRL_OFF(fDistance), pA3dCtrlSuper->fDistance);
	LogMsg(TEXT("...fValue_%x=%g"), A3D_CTRL_OFF(fValue_48), pA3dCtrlSuper->fValue_48);
	LogMsg(TEXT("...sizeof(A3DCTRL_REFLECTION)=%u"), sizeof(A3DCTRL_REFLECTION));

	for (UINT i = 0; i < A3D_MAX_SOURCE_REFLECTIONS; i++)
	{
		for (UINT j = 0; j < sizeof(pA3dCtrlSuper->Reflections[0]); j++)
		{
			LPBYTE pByte = ((LPBYTE)&pA3dCtrlSuper->Reflections[i] + j);
			if (*pByte)
			{
				LogMsg(TEXT("...Reflections[%u]=%#x"), i, A3D_CTRL_OFF(Reflections[i]));
				LogMsg(TEXT("...Reflections[%u].bEnable(%x)=%#x"), i,
					A3D_REFS_OFF(bEnable), pA3dCtrlSuper->Reflections[i].bEnable);
				LogMsg(TEXT("...Reflections[%u].bAvailable(%x)=%#x"), i,
					A3D_REFS_OFF(bAvailable), pA3dCtrlSuper->Reflections[i].bAvailable);
				LogMsg(TEXT("...Reflections[%u].bMute(%x)=%#x"), i,
					A3D_REFS_OFF(bMute), pA3dCtrlSuper->Reflections[i].bMute);
				LogMsg(TEXT("...Reflections[%u].dwValue_%Xh=%#x"), i,
					A3D_REFS_OFF(dwValue_0Ch), pA3dCtrlSuper->Reflections[i].dwValue_0Ch);
				LogMsg(TEXT("...Reflections[%u].fAlpha(%x)=%g"), i,
					A3D_REFS_OFF(fAlpha), pA3dCtrlSuper->Reflections[i].fAlpha);
				LogMsg(TEXT("...Reflections[%u].fLeftEarAzim(%x)=%g"), i,
					A3D_REFS_OFF(LeftEar.fAzim), pA3dCtrlSuper->Reflections[i].LeftEar.fAzim);
				LogMsg(TEXT("...Reflections[%u].fLeftEarElev(%x)=%g"), i,
					A3D_REFS_OFF(LeftEar.fElev), pA3dCtrlSuper->Reflections[i].LeftEar.fElev);
				LogMsg(TEXT("...Reflections[%u].fLeftGain(%x)=%g"), i,
					A3D_REFS_OFF(LeftEar.fGain), pA3dCtrlSuper->Reflections[i].LeftEar.fGain);
				LogMsg(TEXT("...Reflections[%u].fLeftDelay(%x)=%g"), i,
					A3D_REFS_OFF(LeftEar.fDelay), pA3dCtrlSuper->Reflections[i].LeftEar.fDelay);
				LogMsg(TEXT("...Reflections[%u].fRightEarAzim(%x)=%g"), i,
					A3D_REFS_OFF(RightEar.fAzim), pA3dCtrlSuper->Reflections[i].RightEar.fAzim);
				LogMsg(TEXT("...Reflections[%u].fRightEarElev(%x)=%g"), i,
					A3D_REFS_OFF(RightEar.fElev), pA3dCtrlSuper->Reflections[i].RightEar.fElev);
				LogMsg(TEXT("...Reflections[%u].fRightGain(%x)=%g"), i,
					A3D_REFS_OFF(RightEar.fGain), pA3dCtrlSuper->Reflections[i].RightEar.fGain);
				LogMsg(TEXT("...Reflections[%u].fRightDelay(%x)=%g"), i,
					A3D_REFS_OFF(RightEar.fDelay), pA3dCtrlSuper->Reflections[i].RightEar.fDelay);
				LogMsg(TEXT("...Reflections[%u].fAudibility(%x)=%g"), i,
					A3D_REFS_OFF(fAudibility), pA3dCtrlSuper->Reflections[i].fAudibility);
				break;
			}
		}
	}

	LogMsg(TEXT("...fMinDist(%x)=%g"), A3D_CTRL_OFF(fMinDist), pA3dCtrlSuper->fMinDist);
	LogMsg(TEXT("...fMaxDist(%x)=%g"), A3D_CTRL_OFF(fMaxDist), pA3dCtrlSuper->fMaxDist);
	LogMsg(TEXT("...fDistScale(%x)=%g"), A3D_CTRL_OFF(fDistScale), pA3dCtrlSuper->fDistScale);
	//LogMsg(TEXT("...bReserved_3d8=%x"), A3D_CTRL_OFF(bReserved_3d8));

	/*LPDWORD pPrev = 0;
	for (i = 0; i < sizeof(*pA3dCtrlSuper); i++)
	{
		LPBYTE pByte = ((LPBYTE)pA3dCtrlSuper + i);
		if (*pByte)
		{
			UINT j = i & ~3;
			LPDWORD pdw = (LPDWORD)((LPBYTE)pA3dCtrlSuper + j);
			if (pdw != pPrev)
			{
				LogMsg(TEXT("...[%04u/%03x]=%#08x %g"), j, j, *pdw, *(LPA3DVAL)pdw);
				pPrev = pdw;
			}
		}
	}*/
#endif

	// Get DirectSound3DBuffer object for position source sound buffer.
	if (!m_pDS3DB)
	{
		hr = m_pDSB->QueryInterface(IID_IDirectSound3DBuffer, (LPVOID *)&m_pDS3DB);
#ifdef _DEBUG
		LogMsg(TEXT("...QueryInterface(IDirectSound3DBuffer)=%s"), Result(hr));
#endif
		if (FAILED(hr))
			return hr;
	}

	// Calculate average source sound buffer position.
	A3DVAL fAzim = (pA3dCtrlSuper->LeftEar.fAzim + pA3dCtrlSuper->RightEar.fAzim) * 0.5f;
	A3DVAL fElev = (pA3dCtrlSuper->LeftEar.fElev + pA3dCtrlSuper->RightEar.fElev) * 0.5f;

	// Calculate DirectSound3D source sound buffer position.
	DS3DBUFFER DS3DBuffer;
	ZeroMemory(&DS3DBuffer, sizeof(DS3DBuffer));
	DS3DBuffer.dwSize = sizeof(DS3DBuffer);
	DS3DBuffer.vPosition.x = (D3DVALUE)-sin(fAzim);
	DS3DBuffer.vPosition.y = (D3DVALUE)sin(fElev);
	DS3DBuffer.vPosition.z = (D3DVALUE)cos(fAzim);
	DS3DBuffer.dwInsideConeAngle = DS3D_DEFAULTCONEANGLE;
	DS3DBuffer.dwOutsideConeAngle = DS3D_DEFAULTCONEANGLE;
	DS3DBuffer.vConeOrientation.z = 1.0f;
	DS3DBuffer.flMinDistance = 0.5f;
	DS3DBuffer.flMaxDistance = 2.0f;
	DS3DBuffer.dwMode = DS3DMODE_HEADRELATIVE;
#ifdef _DEBUG
	LogMsg(TEXT("...vPosition.x=%g"), DS3DBuffer.vPosition.x);
	LogMsg(TEXT("...vPosition.y=%g"), DS3DBuffer.vPosition.y);
	LogMsg(TEXT("...vPosition.z=%g"), DS3DBuffer.vPosition.z);
#endif

	// Set source sound buffer position.
	hr = m_pDS3DB->SetAllParameters(&DS3DBuffer, DS3D_IMMEDIATE);
#ifdef _DEBUG
	LogMsg(TEXT("...SetAllParameters()=%s"), Result(hr));
#endif
	if (FAILED(hr))
		return hr;

	// Calculate average gain for source sound buffer.
	A3DVAL fGain = (pA3dCtrlSuper->LeftEar.fGain + pA3dCtrlSuper->RightEar.fGain) * 0.5f;

	// Correct gain for equalization effect.
	if (pA3dCtrlSuper->fAlpha > 0.4f)
		fGain *= 1.67f - 1.67f * pA3dCtrlSuper->fAlpha;

	// Calculate volume level for source sound buffer.
	LONG lVolume;
	if (fGain < 0.00001f)
		lVolume = DSBVOLUME_MIN;
	else
		lVolume = (LONG)(log10(fGain) * 2000.0f);

	// Set source sound buffer volume level.
	hr = m_pDSB->SetVolume(lVolume);
#ifdef _DEBUG
	LogMsg(TEXT("...SetVolume(%d)=%s"), lVolume, Result(hr));
#endif
	if (FAILED(hr))
		return hr;

	// Calculate source sound buffer frequency.
	DWORD dwFrequency = (DWORD)(A3D_SAMPLE_RATE_1 * pA3dCtrlSuper->fFreqFactor);

	// Set source sound buffer frequency.
	hr = m_pDSB->SetFrequency(dwFrequency);
#ifdef _DEBUG
	LogMsg(TEXT("...SetFrequency(%u)=%s"), dwFrequency, Result(hr));
#endif
	if (FAILED(hr))
		return hr;

	// Set control packet for reflections.
	if (m_pA3dReflections)
	{
		hr = m_pA3dReflections->SetA3dSuperCtrl(pA3dCtrlSuper, dwFrequency);
#ifdef _DEBUG
		LogMsg(TEXT("...m_pA3dReflections->SetA3dSuperCtrl()=%s"), Result(hr));
#endif
		if (FAILED(hr))
			return hr;
	}

	// Save current source control packet.
	CopyMemory(&m_A3dCtrlSuper, pA3dCtrlSuper, min(dwSize, sizeof(m_A3dCtrlSuper)));

	return S_OK;
}

//===========================================================================
//
// IA3dDalBuffer::GetAllocationStatus
//
// Purpose: Get allocation status for buffer.
//
// Parameters:
//  pdwStatus       LPDWORD pointer to buffer for get allocation status.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dDalBuffer::GetAllocationStatus(LPDWORD pdwStatus)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDalBuffer::GetAllocationStatus(%#x)"), pdwStatus);
	_ASSERTE(pdwStatus && !IsBadWritePtr(pdwStatus, sizeof(*pdwStatus)));
#endif
	// Check arguments values.
	if (!pdwStatus)
		return E_POINTER;

	// Copy allocation status for buffer.
	*pdwStatus = A3D_FALSE;

	return S_OK;
}

//===========================================================================
//
// IA3dDalBuffer::GetWave
//
// Purpose: Get current wave data area.
//
// Parameters:
//  ppbWave         LPDWORD pointer to buffer for get wave data area pointer.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dDalBuffer::GetWave(LPBYTE * ppbWave)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDalBuffer::GetWave(%#x)"), ppbWave);
	_ASSERTE(ppbWave && !IsBadWritePtr(ppbWave, sizeof(*ppbWave)));
#endif
	// Check arguments values.
	if (!ppbWave)
		return E_POINTER;

#if A3D_EMULATE_MAXIMUM
	LPVOID pAudioPtr;
	DWORD dwAudioBytes1, dwAudioBytes2;

	// Lock DirectSoundBuffer data buffer.
	HRESULT hr = m_pDSB->Lock(0, DSBSIZE_MIN, (LPVOID *)ppbWave, &dwAudioBytes1,
		&pAudioPtr, &dwAudioBytes2, 0);
#ifdef _DEBUG
	LogMsg(TEXT("...Lock()=%s"), Result(hr));
#endif
	if (FAILED(hr))
		return hr;

	// Unlock DirectSoundBuffer data buffer.
	hr = m_pDSB->Unlock(*ppbWave, dwAudioBytes1, pAudioPtr, dwAudioBytes2);
#ifdef _DEBUG
	LogMsg(TEXT("...Unlock()=%s"), Result(hr));
#endif
	if (FAILED(hr))
		return hr;

#else
	// Copy wave data area pointer.
	*ppbWave = NULL;
#endif

	return S_OK;
}

//===========================================================================
//
// IA3dDalBuffer::GetDriverInfo
//
// Purpose: Get device driver interfaces.
//
// Parameters:
//  ppIDsDriverBuffer LPVOID* pointer to buffer for DsBuffer interface.
//  ppIA3dDriverBuffer LPVOID* pointer to buffer for A3D driver buffer interface.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dDalBuffer::GetDriverInfo(LPVOID * ppIDsDriverBuffer,
	LPVOID * ppIA3dDriverBuffer)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDalBuffer::GetDriverInfo(%#x,%#x)"), ppIDsDriverBuffer,
		ppIA3dDriverBuffer);
	_ASSERTE(ppIDsDriverBuffer && !IsBadWritePtr(ppIDsDriverBuffer, sizeof(*ppIDsDriverBuffer)));
	_ASSERTE(ppIA3dDriverBuffer && !IsBadWritePtr(ppIA3dDriverBuffer, sizeof(*ppIA3dDriverBuffer)));
#endif
	// Check arguments values.
	if (!ppIDsDriverBuffer || !ppIA3dDriverBuffer)
		return E_POINTER;

	// Copy device driver interfaces.
	*ppIDsDriverBuffer = NULL;
	*ppIA3dDriverBuffer = NULL;

	return S_OK;
}

//===========================================================================
//
// IA3dDalBuffer::SetNewBuffer
//
// Purpose: Set new wave data area.
//
// Parameters:
//  pbWave          LPBYTE pointer to wave data area.
//  dwSize          DWORD wave data area size.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dDalBuffer::SetNewBuffer(LPBYTE pbWave, DWORD dwSize)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDalBuffer::SetNewBuffer(%#x,%u)"), pbWave, dwSize);
	_ASSERTE(pbWave && !IsBadReadPtr(pbWave, sizeof(*pbWave)));
	_ASSERTE(dwSize);
#endif
	// Check arguments values.
	if (!pbWave)
		return E_POINTER;

	// Check arguments values.
	if (!dwSize)
		return E_INVALIDARG;

	// Not implemented.
	return E_NOTIMPL;
}

//===========================================================================
//
// IA3dDalBuffer::SetA3dDirectCtrl
//
// Purpose: Set direct A3D control packet information.
//
// Parameters:
//  pA3dCtrlSuper   LPA3DCTRL_SRC_SUPER pointer to control packet.
//  dwSize          DWORD control packet size.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dDalBuffer::SetA3dDirectCtrl(LPA3DCTRL_SRC_SUPER pA3dCtrlDirect,
	DWORD dwSize)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDalBuffer::SetA3dDirectCtrl(%#x,%u)"), pA3dCtrlDirect, dwSize);
	_ASSERTE(pA3dCtrlDirect && !IsBadReadPtr(pA3dCtrlDirect, sizeof(*pA3dCtrlDirect)));
	_ASSERTE(dwSize);
#endif
	// Check arguments values.
	if (!pA3dCtrlDirect)
		return E_POINTER;

	// Check arguments values.
	if (!dwSize)
		return E_INVALIDARG;

	// Not implemented.
	return E_NOTIMPL;
}


//===========================================================================
//
// IA3dDalBuffer::GetCurrentPositionEx
//
// Purpose: Get extended position information for buffer.
//
// Parameters:
//  pdwPlay         LPDWORD pointer to read position for buffer.
//  pdwWrite        LPDWORD pointer to wtite position for buffer.
//  pdwReflect      LPDWORD pointer to number of reflections for buffer.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dDalBuffer::GetCurrentPositionEx(LPDWORD pdwPlay,
	LPDWORD pdwWrite, LPDWORD pdwReflect)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDalBuffer::GetCurrentPositionEx(%#x,%#x,%#x)"),
		pdwPlay, pdwWrite, pdwReflect);
	_ASSERTE(pdwPlay && !IsBadWritePtr(pdwPlay, sizeof(*pdwPlay)));
	_ASSERTE(pdwWrite && !IsBadWritePtr(pdwWrite, sizeof(*pdwWrite)));
	_ASSERTE(pdwReflect && !IsBadWritePtr(pdwReflect, sizeof(*pdwReflect)));
#endif
	// Check arguments values.
	if (!pdwPlay || !pdwWrite || !pdwReflect)
		return E_POINTER;

	// Copy current positions for buffer.
	*pdwPlay = 0;
	*pdwWrite = 0;

	// Copy number of reflections for buffer.
	if (m_pA3dReflections)
		*pdwReflect = m_pA3dReflections->ReadyForService();
	else
		*pdwReflect = 0;

	return S_OK;
}


//===========================================================================
//
// IA3dDalBuffer::GetStatusEx
//
// Purpose: Get extended status for buffer.
//
// Parameters:
//  pdwStatus       LPDWORD pointer to extended status for buffer.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dDalBuffer::GetStatusEx(LPDWORD pdwStatus)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dDalBuffer::GetStatusEx(%#x)"), pdwStatus);
	_ASSERTE(pdwStatus && !IsBadWritePtr(pdwStatus, sizeof(*pdwStatus)));
#endif
	// Check arguments values.
	if (!pdwStatus)
		return E_POINTER;

	// Copy extended status for buffer.
	*pdwStatus = 0;

	return S_OK;
}
