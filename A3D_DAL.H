//===========================================================================
//
// A3D_DAL.H
//
// Purpose: Emulate Aureal A3D DAL interface (Wrapper to DirectSound3D).
// 
// Copyright (C) 2004 Dmitry Nesterenko. All rights reserved.
//
//===========================================================================


#ifndef _A3D_DAL_H_
#define _A3D_DAL_H_


//===========================================================================
//
// Class and interface ID's for A3D DAL interfaces.
//
//===========================================================================

// A3D DAL class ID: {442D12A1-2641-11D2-90FB-006008A1F441}
DEFINE_GUID(CLSID_A3dDal, 0x442d12a1, 0x2641, 0x11d2,
	0x90, 0xfb, 0x0, 0x60, 0x8, 0xa1, 0xf4, 0x41);

// A3dScaleHack interface ID: {D77CFAAC-CA67-11D2-BB49-0060082F3C00}
DEFINE_GUID(IID_IA3dScaleHack, 0xd77cfaac, 0xca67, 0x11d2,
	0xbb, 0x49, 0x0, 0x60, 0x8, 0x2f, 0x3c, 0x0);

// A3dScaleHackBuffer interface ID: {B4279161-CB06-11D2-BB49-0060082F3C00})	
DEFINE_GUID(IID_IA3dScaleHackBuffer, 0xb4279161, 0xcb06, 0x11d2,
	0xbb, 0x49, 0x0, 0x60, 0x8, 0x2f, 0x3c, 0x0);

// A3dDal interface ID: {A3C62F2A-2671-11D2-BB48-0060082F3C00}
DEFINE_GUID(IID_IA3dDal, 0xa3c62f2a, 0x2671, 0x11d2,
	0xbb, 0x48, 0x0, 0x60, 0x8, 0x2f, 0x3c, 0x0);

// A3dDalBuffer interface ID: {E04814D8-26FC-11D2-BB48-0060082F3C00}
DEFINE_GUID(IID_IA3dDalBuffer, 0xe04814d8, 0x26fc, 0x11d2,
	0xbb, 0x48, 0x0, 0x60, 0x8, 0x2f, 0x3c, 0x0);

// A3dDalBuffer2 interface ID: {C9D08FE1-646C-11D2-BE63-006097CECF9F}
DEFINE_GUID(IID_IA3dDalBuffer2, 0xc9d08fe1, 0x646c, 0x11d2,
	0xbe, 0x63, 0x0, 0x60, 0x97, 0xce, 0xcf, 0x9f);

// A3D DAL capabilities property set ID: {08A67509-3842-11D3-BB49-0060082F3C00}
DEFINE_GUID(DSPROPERTY_A3D_GET_CAPS, 0x8a67509, 0x3842, 0x11d3,
	0xbb, 0x49, 0x0, 0x60, 0x8, 0x2f, 0x3c, 0x0);

// A3D DAL set reflection property set ID: {3F9E77A1-3852-11D3-BB49-0060082F3C00}
DEFINE_GUID(DSPROPERTY_A3DBUFFER_SET_REFLECTION, 0x3f9e77a1, 0x3852, 0x11d3,
	0xbb, 0x49, 0x0, 0x60, 0x8, 0x2f, 0x3c, 0x0);


//===========================================================================
//
// Forward class declarations for A3D DAL interfaces.
// 
//===========================================================================
class IA3dScaleHack;
class IA3dScaleHackBuffer;
class IA3dDal;
class IA3dDalBuffer;
class IA3dReflections;

typedef class IA3dScaleHack			*LPA3DSCALEHACK;
typedef class IA3dScaleHackBuffer		*LPA3DSCALEHACKBUFFER;
typedef class IA3dDal				*LPA3DDAL;
typedef class IA3dDalBuffer			*LPA3DDALBUFFER;
typedef class IA3dReflections			*LPA3DREFLECTIONS;

//===========================================================================
//
// Defined values for A3D DAL interfaces.
// 
//===========================================================================

// Emulate maximum compatibility for A3D DAL calls.
#define A3D_EMULATE_MAXIMUM			TRUE

// Sample rate definitions.
#define A3D_SAMPLE_RATE_1			22050
#define A3D_SAMPLE_RATE_2			44100

// Initialize render flags.
#define A3DINIT_DEFAULT1			0x00000001
#define A3DINIT_DEFAULT2			0x00000008
#define A3DINIT_DISABLE_SPLASHSCREEN		0x00000002
#define A3DINIT_DISABLE_DS3D			0x00000010
#define A3DINIT_DISABLE_DS			0x00000020
#define A3DINIT_DISABLE_A3D			0x00000040

// Hardware device types.
#define A3DCAPS_DEVICE_TYPE_AU8810		0x010B
#define A3DCAPS_DEVICE_TYPE_AU8820		0x0101
#define A3DCAPS_DEVICE_TYPE_AU8830		0x0104
#define A3DCAPS_DEVICE_TYPE_A2D			0x010C
#define A3DCAPS_DEVICE_TYPE_D3D			0x010D

// Version definitions for A3D_CAPS.
#define A3DCAPS_VERSION				0x0110

// Factor for calculate from float value to fixed.
#define A3DCAPS_CALC_FACTOR			32767

// Constants for defferent devices.
#define A3DCAPS_1CH_AU8810			0x0823
#define A3DCAPS_1CH_AU8820			0x0823
#define A3DCAPS_1CH_AU8830			0x000E
#define A3DCAPS_1CH_EMULATE			0x000C

#define A3DCAPS_44H_AU8810			0x04060C94
#define A3DCAPS_44H_AU8820			0x040604A5
#define A3DCAPS_44H_AU8830			0x04060800
#define A3DCAPS_44H_AU8810_WDM			0x050C0DAC
#define A3DCAPS_44H_AU8820_WDM			0x050A05DC
#define A3DCAPS_44H_AU8830_WDM			0x050C09C4

// Feature flags for capabilities.
#define A3DCAPS_AU8820_DEFAULT			0x00000033
#define A3DCAPS_AU8830_DEFAULT			0x00000117
#define A3DCAPS_ALL_PRIORITY			0x00001000
#define A3DCAPS_NO_SPLASHSCREEN			0x00002000
#define A3DCAPS_USE_DS_BUFFERS			0x00010000
#define A3DCAPS_NEED_SCALE_HACK			0x00020000
#define A3DCAPS_IS_EMU_DRIVER			0x00040000
#define A3DCAPS_IS_WDM_DRIVER			0x00080000

// Maximum number of reflections for one source.
#define A3D_MAX_SOURCE_REFLECTIONS		16

// Constants for sample capabilities.
#define A3DCAPS_SAMPLE_1AH			0x0100
#define A3DCAPS_SAMPLE_1CH			0x0FFF

// HTTF mode definitions.
#define A3DCAPS_A2DHRTF				1
#define A3DCAPS_HAWK13				13
#define A3DCAPS_HAWK22				25
#define A3DCAPS_HAWK44				50

// Maximum number of samples.
#define A3DCAPS_MAX_SAMPLES			8

// Request functionality definitions.
enum
{
	A3DFUNC_0,
	A3DFUNC_1,
	A3DFUNC_2,
	A3DFUNC_3,
	A3DFUNC_4,
	A3DFUNC_NOT_EXCLUSIVE,
	A3DFUNC_6,
	A3DFUNC_NO_SPLASHSCREEN,
	A3DFUNC_8,
	A3DFUNC_9,
	A3DFUNC_A,
	A3DFUNC_B,
	A3DFUNC_C,
	A3DFUNC_D,
	A3DFUNC_USE_DS_BUFFERS,
	A3DFUNC_NEED_SCALE_HACK
};


//===========================================================================
//
// Structures for A3D DAL interfaces.
// 
//===========================================================================

// Structure information about DirectSound driver.
typedef struct __DSDRIVERDESC
{
	DWORD dwFlags;				// Characteristics of the driver.
	CHAR szDesc[256];			// Description of the driver.
	CHAR szDrvname[256];			// File name of the driver module.
	DWORD dnDevNode;			// Devnode of this device.
	WORD wVxdId;				// VxD ID for this driver.
	WORD wReserved;
	ULONG ulDeviceNum;			// Wave device ID used by this driver.
	DWORD dwHeapType;			// Type of heap onboard sound memory.
	LPVOID pvDirectDrawHeap;		// DirectDraw heap pointer.
	DWORD dwMemStartAddress;		// Start addresse onboard sound memory.
	DWORD dwMemEndAddress;			// End address onboard sound memory.
	DWORD dwMemAllocExtra;			// Number of extra bytes in sound buffer.
	LPVOID pvReserved1;
	LPVOID pvReserved2;
} DSDRIVERDESC, *LPDSDRIVERDESC;

// Structure information about sample capabilities.
typedef struct __A3D_SAMPLE_CAPS
{
	WORD wEnable;				// Sample used or no.
	BYTE unknown_2[2];
	DWORD dwSamplesPerSec1;			// Samples per second.
	DWORD dwSamplesPerSec2;			// -"-
	WORD wCtrlSize;				// Size of control structure.
	WORD wCtrlOffsetLeft;			// Offset in control structure.
	WORD wValue_10h;			// 0?
	WORD wCtrlOffsetRight;			// Offset in control structure.
	WORD wValue_14h;			// 0?
	WORD wHrtfMode;				// 3D HRTF mode.
	WORD wBitsPerSample;			// Bits per sample.
	WORD wValue_1Ah;			// 0xfff?
	WORD wValue_1Ch;			// 0x100,0x240?
	BYTE unknown_1Eh[2];
	DWORD dwValue_20h;			// ?
	BYTE unknown_24h[20];
} A3D_SAMPLE_CAPS, *LPA3D_SAMPLE_CAPS;

// Structure information about A3D capabilities.
typedef struct __A3D_CAPS
{
	WORD wDeviceType;			// Hardware device type.
	WORD wVersion;				// A3D DAL interface version in bcd.
	DWORD dwCalcFactor;			// Always 32767?
	WORD wValue_8;				// Hardware channels = 2,4?
	WORD wChannels;				// Number of channals.
	WORD wRefMask;				// For reflections = 0,7?
	BYTE unknown_0Eh[2];
	DWORD dwMaxSampleRate;			// Number samples per second.
	WORD wMaxBuffers;			// Number of sound buffers.
	BYTE unknown_16h[2];
	DWORD dwSamplesMask;			// Bit mask for samples capabilities.
	DWORD dwValue_1Ch;			// Flags = 0xc,0xe,0x823?
	DWORD dwFeatureFlags;			// Flags features = 0x90117?
	BYTE unknown_24h[4];
	DWORD dwMaxReflections;			// Total number supported reflections.
	DWORD dwMaxSrcReflections;		// Maximum reflections for one source.
	DWORD dwValue_30h;			// For reflections = 0?
	DWORD dwValue_34h;			// For reflections = ?
	DWORD dwValue_38h;			// For reflections = ?
	BYTE unknown_3Ch[8];
	DWORD dwValue_44h;			// Flags?
	BYTE unknown_48h[4];
	A3D_SAMPLE_CAPS Samples[A3DCAPS_MAX_SAMPLES];
	BYTE unknown_20Ch[40];
} A3D_CAPS, *LPA3D_CAPS;

// Structure information forcontrol channel property.
typedef struct __A3DCTRL_EAR
{
	A3DVAL fAzim;				// Sound direction azimuth.
	A3DVAL fElev;				// Sound direction elevation.
	A3DVAL fGain;				// Sound volume gain.
	A3DVAL fDelay;				// Sound delay.
} A3DCTRL_EAR, *LPA3DCTRL_EAR;

// Structure information for control reflection property.
typedef struct __A3DCTRL_REFLECTION
{
	BOOL bEnable;				// Enable?
	BOOL bAvailable;			// Available?
	BOOL bMute;				// Mute?
	DWORD dwValue_0Ch;			// (0)?
	A3DVAL fAlpha;				// Equalization effect.
	A3DCTRL_EAR LeftEar;			// Left channel.
	A3DCTRL_EAR RightEar;			// Right channel.
	A3DVAL fAudibility;			// Calculated audibility.
} A3DCTRL_REFLECTION, *LPA3DCTRL_REFLECTION;

// Structure information for control buffer property.
typedef struct __A3DCTRL_SRC_SUPER
{
	DWORD dwValue_0;			// (0)?
	DWORD dwMode;				// Source mode (1,Mono=2).
	BOOL bExecuted;				// Executed?
	A3DVAL fPitch;				// Pitch?
	A3DVAL fAlpha;				// Equalization effect.
	A3DVAL fFreqFactor;			// Change playback frequency.
	A3DVAL fNative;				// (0,Mono+Native=1).
	A3DCTRL_EAR LeftEar;			// Left channel.
	A3DCTRL_EAR RightEar;			// Right channel.
	A3DVAL fPriority;			// Buffer priority.
	A3DVAL fAudibility;			// Calculated audibility.
	A3DVAL fDistance;			// Distance.
	A3DVAL fValue_48;			// ?
	A3DCTRL_REFLECTION Reflections[A3D_MAX_SOURCE_REFLECTIONS];
	A3DVAL fMinDist;			// Minimum distance.
	A3DVAL fMaxDist;			// Maximum distance.
	A3DVAL fDistScale;			// Distance model scale.
	BYTE unknown_3D8h[52];
} A3DCTRL_SRC_SUPER, *LPA3DCTRL_SRC_SUPER;


//===========================================================================
//
// This class is the A3dScaleHack objects.
// 
//===========================================================================
class IA3dScaleHack : public IUnknown
{
protected:
	// IA3dScaleHack internal members.
	LONG m_cRef;
	BOOL m_bScaleHack;

public:
	// Constructor and destructor.
	IA3dScaleHack();
	~IA3dScaleHack();

	// IUnknown members.
	STDMETHOD(QueryInterface)(REFIID, LPVOID *);
	STDMETHOD_(ULONG, AddRef)();
	STDMETHOD_(ULONG, Release)();

	// IA3dScaleHack methods.
	STDMETHOD(ScaleHackEnable)();
	STDMETHOD(ScaleHackDisable)();
	STDMETHOD(IsScaleHackNeeded)(LPBOOL);
};


//===========================================================================
//
// This class is the A3dScaleHackBuffer objects.
// 
//===========================================================================
class IA3dScaleHackBuffer : public IUnknown
{
protected:
	// IA3dScaleHackBuffer internal members.
	LONG m_cRef;
	BOOL m_bScaleHackBuffer;

public:
	// Constructor and destructor.
	IA3dScaleHackBuffer();
	~IA3dScaleHackBuffer();

	// IUnknown members.
	STDMETHOD(QueryInterface)(REFIID, LPVOID *);
	STDMETHOD_(ULONG, AddRef)();
	STDMETHOD_(ULONG, Release)();

	// IA3dScaleHackBuffer methods.
	STDMETHOD(SetScaleHack)(BOOL);
};


//===========================================================================
//
// This class is the A3dDal objects.
// 
//===========================================================================
class IA3dDal : public IUnknown
{
protected:
	// IA3dDal internal members.
	STDMETHODIMP SetA3dCaps();

	LONG m_cRef;
	A3D_CAPS m_A3dCaps;
	LPDIRECTSOUND m_pDS;

public:
	// Constructor and destructor.
	IA3dDal(LPDIRECTSOUND);
	~IA3dDal();

	// IUnknown members.
	STDMETHOD(QueryInterface)(REFIID, LPVOID *);
	STDMETHOD_(ULONG, AddRef)();
	STDMETHOD_(ULONG, Release)();

	// IA3dDal methods.
	STDMETHOD(InitializeEx)(LPGUID, DWORD, DWORD, LPDWORD);
	STDMETHOD(CreateSoundBufferEx)(LPCDSBUFFERDESC, LPBYTE, LPDIRECTSOUNDBUFFER *, LPUNKNOWN);
	STDMETHOD(GetA3dCaps)(LPA3D_CAPS, LPDWORD);
	STDMETHOD(GetDriverInfo)(LPHANDLE, LPVOID *, LPVOID *, LPDWORD);
	STDMETHOD(GetDS)(LPDIRECTSOUND *);
	STDMETHOD(GetDSDriverDesc)(LPDSDRIVERDESC, LPDWORD);
	STDMETHOD(QueryFunctionality)(DWORD, LPDWORD);
	STDMETHOD(Verify)(LPSTR, LPSTR *, LPSTR *);
};


//===========================================================================
//
// This class is the A3dDalBuffer objects.
// 
//===========================================================================
class IA3dDalBuffer : public IUnknown
{
protected:
	LONG m_cRef;
	A3DCTRL_SRC_SUPER m_A3dCtrlSuper;
	LPDIRECTSOUND m_pDS;
	LPDIRECTSOUNDBUFFER m_pDSB;
	LPDIRECTSOUND3DBUFFER m_pDS3DB;
	LPA3DREFLECTIONS m_pA3dReflections;

public:
	// Constructor and destructor.
	IA3dDalBuffer(LPDIRECTSOUND, LPDIRECTSOUNDBUFFER);
	~IA3dDalBuffer();

	// IUnknown members.
	STDMETHOD(QueryInterface)(REFIID, LPVOID *);
	STDMETHOD_(ULONG, AddRef)();
	STDMETHOD_(ULONG, Release)();

	// IA3dDalBuffer methods.
	STDMETHOD(SetA3dSuperCtrl)(LPA3DCTRL_SRC_SUPER, DWORD);
	STDMETHOD(GetAllocationStatus)(LPDWORD);
	STDMETHOD(GetWave)(LPBYTE *);
	STDMETHOD(GetDriverInfo)(LPVOID *, LPVOID *);
	STDMETHOD(SetNewBuffer)(LPBYTE, DWORD);
	STDMETHOD(SetA3dDirectCtrl)(LPA3DCTRL_SRC_SUPER, DWORD);

	// IA3dDalBuffer2 methods.
	STDMETHOD(GetCurrentPositionEx)(LPDWORD, LPDWORD, LPDWORD);
	STDMETHOD(GetStatusEx)(LPDWORD);
};


#endif // _A3D_DAL_H_
