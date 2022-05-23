//===========================================================================
//
// A3D_REF.H
//
// Purpose: Work with A3D 1st reflections (Wrapper to DirectSound3D).
// 
// Copyright (C) 2004 Dmitry Nesterenko. All rights reserved.
//
//===========================================================================


#ifndef _A3D_REF_H_
#define _A3D_REF_H_


//===========================================================================
//
// Service thread function for play 1st reflections.
//
//===========================================================================
DWORD WINAPI ServiceThread(LPVOID);


//===========================================================================
//
// Forward class declarations for A3D 1st reflections.
// 
//===========================================================================
class IA3dReflections;

typedef class IA3dReflections		*LPA3DREFLECTIONS;


//===========================================================================
//
// Defined values for A3D 1st reflections.
// 
//===========================================================================

// Maximal waiting time of end of service thread (msec).
#define A3DREF_MAX_WAIT_THREAD		10

// Precision of reflection delay (~msec).
#define A3DREF_DELAY_PRECISION		5

// Change of frequency for compensation delay (x/128).
#define A3DREF_CHANGE_FREQUENCY		4


//===========================================================================
//
// This class is the A3dReflections objects.
// 
//===========================================================================
class IA3dReflections
{
protected:
	// IA3dReflections friend functions.
	friend DWORD WINAPI ServiceThread(LPVOID);

	// IA3dReflections internal members.
	STDMETHODIMP CreateReflection(DWORD);
	STDMETHODIMP SchedulePlay(DWORD);
	STDMETHODIMP PlayWithLag(DWORD, DWORD);
	STDMETHODIMP_(VOID) Service();
	STDMETHODIMP_(VOID) Reset(DWORD);
	STDMETHODIMP_(VOID) Stop();

	DWORD m_dwBufferSize;
	DWORD m_dwBytesPerSample;
	DWORD m_dwSourceFrequency;
	LPDIRECTSOUND m_pDS;
	LPDIRECTSOUNDBUFFER m_pDSB;
	LPDIRECTSOUNDNOTIFY m_pDSN;
	LPDIRECTSOUNDBUFFER m_pRefsDSB[A3D_MAX_SOURCE_REFLECTIONS];
	LPDIRECTSOUND3DBUFFER m_pRefsDS3DB[A3D_MAX_SOURCE_REFLECTIONS];
	DSBPOSITIONNOTIFY m_DSBPN[A3D_MAX_SOURCE_REFLECTIONS + 1];
	DS3DBUFFER m_DS3DBuffer;
	CRITICAL_SECTION m_CS;
	HANDLE m_hThread;

public:
	// Constructor and destructor.
	IA3dReflections();
	~IA3dReflections();

	// IA3dReflections methods.
	STDMETHODIMP Initialize(LPDIRECTSOUND, LPDIRECTSOUNDBUFFER);
	STDMETHODIMP SetA3dSuperCtrl(LPA3DCTRL_SRC_SUPER, DWORD);
	STDMETHODIMP TrackDelay();
	STDMETHODIMP_(DWORD) ReadyForService();
};


#endif // _A3D_REF_H_
