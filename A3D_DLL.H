//===========================================================================
//
// A3D_DLL.H
//
// Purpose: Aureal A3D 1.x interface library (Wrapper to DirectSound3D).
//
// Copyright (C) 2004 Dmitry Nesterenko. All rights reserved.
//
//===========================================================================


#ifndef _A3D_DLL_H_
#define _A3D_DLL_H_


//===========================================================================
// 
// Internal global variables for A3D library module.
// 
//===========================================================================
extern HMODULE g_hModule;
extern LONG g_cLock;
extern LONG g_cObj;


//===========================================================================
// 
// Internal general utility functions for A3D library module.
// 
//===========================================================================
#ifdef _DEBUG
void LogMsg(LPCTSTR, ...);
LPCTSTR Result(HRESULT);
#endif
LPTSTR GuidToStr(REFGUID, LPTSTR, UINT);


//===========================================================================
// 
// External entry point for A3D library module.
// 
//===========================================================================
extern "C" HRESULT WINAPI A3dCreate(LPGUID, LPDIRECTSOUND *, LPUNKNOWN);


//===========================================================================
//
// Class and interface ID's for A3D 1.x interfaces.
//
//===========================================================================

// ResManMode property set ID (for Creative API only): {62A69BAE-DF9D-11D1-99A6-00C04FC99D46}
DEFINE_GUID(DSPROPSETID_RMM_Creative, 0x62a69bae, 0xdf9d, 0x11d1,
	0x99, 0xa6, 0x0, 0xc0, 0x4f, 0xc9, 0x9d, 0x46);

// ResManMode property set ID (for Sensaura API only): {BCD71AE0-CD70-11D1-87F3-0020AFC95782}
DEFINE_GUID(DSPROPSETID_RMM_Sensaura, 0xbcd71ae0, 0xcd70, 0x11d1,
	0x87, 0xf3, 0x0, 0x20, 0xaf, 0xc9, 0x57, 0x82);

// HFAbsorbFactor property set ID (for Sensaura API only): {24A4A6C0-900D-11D1-B9D1-D9D09BDE7E20}
DEFINE_GUID(DSPROPSETID_HFAF_Sensaura, 0x24a4a6c0, 0x900d, 0x11d1,
	0xb9, 0xd1, 0xd9, 0xd0, 0x9b, 0xde, 0x7e, 0x20);


//===========================================================================
//
// Forward class declarations for A3D 1.x interfaces.
// 
//===========================================================================
class IA3dClassFactory;
class IA3dX;
class IA3dSound;
class IA3dSoundBuffer;
class IA3dSound3DListener;
class IA3dSound3DBuffer;
class IA3dKsPropertySet;

typedef class IA3dClassFactory		*LPA3DCLASSFACTORY;
typedef class IA3dX			*LPA3DX;
typedef class IA3dSound			*LPA3DSOUND;
typedef class IA3dSoundBuffer		*LPA3DSOUNDBUFFER;
typedef class IA3dSound3DListener	*LPA3DSOUND3DLISTENER;
typedef class IA3dSound3DBuffer		*LPA3DSOUND3DBUFFER;
typedef class IA3dKsPropertySet		*LPA3DKSPROPERTYSET;


//===========================================================================
//
// Defined values for A3D 1.x interfaces.
// 
//===========================================================================

// Minimum number of available hardware sound buffers.
#define A3D_MIN_HARDWARE_BUFFERS	8

// Feature flags for A3D capabilities.
#define A3D_SOFTWARECAPS_DEFAULT	0x00000005
#define A3D_HARDWARECAPS_DEFAULT	0x00000010


//===========================================================================
// 
// This class is the class factory for all A3D objects.
// 
//===========================================================================
class IA3dClassFactory : public IClassFactory
{
protected:
	// IA3dClassFactory internal members.
	LONG m_cRef;

public:
	// Constructor and destructor.
	IA3dClassFactory();
	~IA3dClassFactory();

	// IUnknown members.
	STDMETHOD(QueryInterface)(REFIID, LPVOID *);
	STDMETHOD_(ULONG, AddRef)();
	STDMETHOD_(ULONG, Release)();

	// IClassFactory members.
	STDMETHOD(CreateInstance)(LPUNKNOWN, REFIID, LPVOID *);
	STDMETHOD(LockServer)(BOOL);
};


//===========================================================================
// 
// This class is the A3dX objects.
// 
//===========================================================================
class IA3dX : public IA3d2
{
protected:
	// IA3dX internal members.
	STDMETHODIMP CreateKsPropertySet();
	STDMETHODIMP SetOemResourceManagerMode(DWORD);
	STDMETHODIMP GetOemResourceManagerMode(LPDWORD);
	STDMETHODIMP SetOemHFAbsorbFactor(FLOAT);

	LONG m_cRef;
	DWORD m_dwAppVersion;
	DWORD m_dwResourceManagerMode;
	DWORD m_dwFrontXtalkMode;
	DWORD m_dwBackXtalkMode;
	DWORD m_dwQuadMode;
	FLOAT m_fHFAbsorbFactor;
	LPDIRECTSOUND m_pDS;
	LPKSPROPERTYSET m_pKsPS;

public:
	// Constructor and destructor.
	IA3dX(LPDIRECTSOUND);
	~IA3dX();

	// IUnknown members.
	STDMETHOD(QueryInterface)(REFIID, LPVOID *);
	STDMETHOD_(ULONG, AddRef)();
	STDMETHOD_(ULONG, Release)();

	// IA3d methods.
	STDMETHOD(SetOutputMode)(DWORD, DWORD, DWORD);
	STDMETHOD(GetOutputMode)(LPDWORD, LPDWORD, LPDWORD);
	STDMETHOD(SetResourceManagerMode)(DWORD);
	STDMETHOD(GetResourceManagerMode)(LPDWORD);
	STDMETHOD(SetHFAbsorbFactor)(FLOAT);
	STDMETHOD(GetHFAbsorbFactor)(FLOAT *);

	// IA3d2 methods.
	STDMETHOD(RegisterVersion)(DWORD);
	STDMETHOD(GetSoftwareCaps)(LPA3DCAPS_SOFTWARE);
	STDMETHOD(GetHardwareCaps)(LPA3DCAPS_HARDWARE);
};

typedef IA3dX *LPA3DX;


//===========================================================================
// 
// This class is the A3dSound objects.
// 
//===========================================================================
class IA3dSound : public IDirectSound
{
protected:
	// IA3dSound internal members.
	LONG m_cRef;
	LPDIRECTSOUND m_pDS;

public:
	// Constructor and destructor.
	IA3dSound(LPDIRECTSOUND);
	~IA3dSound();

	// IUnknown members.
	STDMETHOD(QueryInterface)(REFIID, LPVOID *);
	STDMETHOD_(ULONG, AddRef)();
	STDMETHOD_(ULONG, Release)();
    
	// IDirectSound methods.
	STDMETHOD(CreateSoundBuffer)(LPCDSBUFFERDESC, LPDIRECTSOUNDBUFFER *, LPUNKNOWN);
	STDMETHOD(GetCaps)(LPDSCAPS);
	STDMETHOD(DuplicateSoundBuffer)(LPDIRECTSOUNDBUFFER, LPDIRECTSOUNDBUFFER *);
	STDMETHOD(SetCooperativeLevel)(HWND, DWORD);
	STDMETHOD(Compact)();
	STDMETHOD(GetSpeakerConfig)(LPDWORD);
	STDMETHOD(SetSpeakerConfig)(DWORD);
	STDMETHOD(Initialize)(LPCGUID);
};


//===========================================================================
// 
// This class is the A3dSoundBuffer objects.
// 
//===========================================================================
class IA3dSoundBuffer : public IDirectSoundBuffer
{
protected:
	// IA3dSoundBuffer internal members.
	LONG m_cRef;
	LPDIRECTSOUNDBUFFER m_pDSB;
	LPDIRECTSOUND m_pDS;

public:
	// Constructor and destructor.
	IA3dSoundBuffer(LPDIRECTSOUNDBUFFER, LPDIRECTSOUND);
	~IA3dSoundBuffer();

	// IUnknown members.
	STDMETHOD(QueryInterface)(REFIID, LPVOID *);
	STDMETHOD_(ULONG, AddRef)();
	STDMETHOD_(ULONG, Release)();

	// IDirectSoundBuffer methods
	STDMETHOD(GetCaps)(LPDSBCAPS);
	STDMETHOD(GetCurrentPosition)(LPDWORD, LPDWORD);
	STDMETHOD(GetFormat)(LPWAVEFORMATEX, DWORD, LPDWORD);
	STDMETHOD(GetVolume)(LPLONG);
	STDMETHOD(GetPan)(LPLONG);
	STDMETHOD(GetFrequency)(LPDWORD);
	STDMETHOD(GetStatus)(LPDWORD);
	STDMETHOD(Initialize)(LPDIRECTSOUND, LPCDSBUFFERDESC);
	STDMETHOD(Lock)(DWORD, DWORD, LPVOID *, LPDWORD, LPVOID *, LPDWORD, DWORD);
	STDMETHOD(Play)(DWORD, DWORD, DWORD);
	STDMETHOD(SetCurrentPosition)(DWORD);
	STDMETHOD(SetFormat)(LPCWAVEFORMATEX);
	STDMETHOD(SetVolume)(LONG);
	STDMETHOD(SetPan)(LONG);
	STDMETHOD(SetFrequency)(DWORD);
	STDMETHOD(Stop)();
	STDMETHOD(Unlock)(LPVOID, DWORD, LPVOID, DWORD);
	STDMETHOD(Restore)();
};


//===========================================================================
// 
// This class is the A3dSound3DListener objects.
// 
//===========================================================================
class IA3dSound3DListener : public IDirectSound3DListener
{
protected:
	// IA3dSound3DListener internal members.
	LONG m_cRef;
	LPDIRECTSOUND3DLISTENER m_pDS3DL;

public:
	// Constructor and destructor.
	IA3dSound3DListener(LPDIRECTSOUND3DLISTENER);
	~IA3dSound3DListener();

	// IUnknown members.
	STDMETHOD(QueryInterface)(REFIID, LPVOID *);
	STDMETHOD_(ULONG, AddRef)();
	STDMETHOD_(ULONG, Release)();

	// IDirectSound3DListener methods
	STDMETHOD(GetAllParameters)(LPDS3DLISTENER);
	STDMETHOD(GetDistanceFactor)(LPD3DVALUE);
	STDMETHOD(GetDopplerFactor)(LPD3DVALUE);
	STDMETHOD(GetOrientation)(LPD3DVECTOR, LPD3DVECTOR);
	STDMETHOD(GetPosition)(LPD3DVECTOR);
	STDMETHOD(GetRolloffFactor)(LPD3DVALUE);
	STDMETHOD(GetVelocity)(LPD3DVECTOR);
	STDMETHOD(SetAllParameters)(LPCDS3DLISTENER, DWORD);
	STDMETHOD(SetDistanceFactor)(D3DVALUE, DWORD);
	STDMETHOD(SetDopplerFactor)(D3DVALUE, DWORD);
	STDMETHOD(SetOrientation)(D3DVALUE, D3DVALUE, D3DVALUE, D3DVALUE, D3DVALUE, D3DVALUE, DWORD);
	STDMETHOD(SetPosition)(D3DVALUE, D3DVALUE, D3DVALUE, DWORD);
	STDMETHOD(SetRolloffFactor)(D3DVALUE, DWORD);
	STDMETHOD(SetVelocity)(D3DVALUE, D3DVALUE, D3DVALUE, DWORD);
	STDMETHOD(CommitDeferredSettings)();
};


//===========================================================================
// 
// This class is the A3dSound3DBuffer objects.
// 
//===========================================================================
class IA3dSound3DBuffer : public IDirectSound3DBuffer
{
protected:
	// IA3dSound3DBuffer internal members.
	LONG m_cRef;
	LPDIRECTSOUND3DBUFFER m_pDS3DB;

public:
	// Constructor and destructor.
	IA3dSound3DBuffer(LPDIRECTSOUND3DBUFFER);
	~IA3dSound3DBuffer();

	// IUnknown members.
	STDMETHOD(QueryInterface)(REFIID, LPVOID *);
	STDMETHOD_(ULONG, AddRef)();
	STDMETHOD_(ULONG, Release)();

	// IDirectSound3DBuffer methods
	STDMETHOD(GetAllParameters)(LPDS3DBUFFER);
	STDMETHOD(GetConeAngles)(LPDWORD, LPDWORD);
	STDMETHOD(GetConeOrientation)(LPD3DVECTOR);
	STDMETHOD(GetConeOutsideVolume)(LPLONG);
	STDMETHOD(GetMaxDistance)(LPD3DVALUE);
	STDMETHOD(GetMinDistance)(LPD3DVALUE);
	STDMETHOD(GetMode)(LPDWORD);
	STDMETHOD(GetPosition)(LPD3DVECTOR);
	STDMETHOD(GetVelocity)(LPD3DVECTOR);
	STDMETHOD(SetAllParameters)(LPCDS3DBUFFER, DWORD);
	STDMETHOD(SetConeAngles)(DWORD, DWORD, DWORD);
	STDMETHOD(SetConeOrientation)(D3DVALUE, D3DVALUE, D3DVALUE, DWORD);
	STDMETHOD(SetConeOutsideVolume)(LONG, DWORD);
	STDMETHOD(SetMaxDistance)(D3DVALUE, DWORD);
	STDMETHOD(SetMinDistance)(D3DVALUE, DWORD);
	STDMETHOD(SetMode)(DWORD, DWORD);
	STDMETHOD(SetPosition)(D3DVALUE, D3DVALUE, D3DVALUE, DWORD);
	STDMETHOD(SetVelocity)(D3DVALUE, D3DVALUE, D3DVALUE, DWORD);
};


//===========================================================================
// 
// This class is the A3dKsPropertySet objects.
// 
//===========================================================================
class IA3dKsPropertySet : public IKsPropertySet
{
protected:
	// IA3dKsPropertySet internal members.
	STDMETHODIMP GetA3dCaps(LPVOID, LPDWORD);

	LONG m_cRef;
	LPKSPROPERTYSET m_pKsPS;

public:
	// Constructor and destructor.
	IA3dKsPropertySet(LPKSPROPERTYSET);
	~IA3dKsPropertySet();

	// IUnknown members.
	STDMETHOD(QueryInterface)(REFIID, LPVOID *);
	STDMETHOD_(ULONG, AddRef)();
	STDMETHOD_(ULONG, Release)();

	// IKsPropertySet methods
	STDMETHOD(Get)(REFGUID, ULONG, LPVOID, ULONG, LPVOID, ULONG, PULONG);
	STDMETHOD(Set)(REFGUID, ULONG, LPVOID, ULONG, LPVOID, ULONG);
	STDMETHOD(QuerySupport)(REFGUID, ULONG, PULONG);
};


#endif // _A3D_DLL_H_
