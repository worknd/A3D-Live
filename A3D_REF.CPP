//===========================================================================
//
// A3D_REF.CPP
//
// Purpose: Work with 1st reflections (Wrapper to DirectSound3D).
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


//===========================================================================
//
// ::ServiceThread
//
// Purpose: Callback function for service thread.
//
// Parameters:
//  pA3dReflections LPVOID pointer to A3dReflections object.
//
// Return: NO_ERROR always.
//
//===========================================================================
DWORD WINAPI ServiceThread(LPVOID pA3dReflections)
{
#ifdef _DEBUG
	LogMsg(TEXT("ServiceThread(%#x)"), pA3dReflections);
	_ASSERTE(pA3dReflections);
#endif
	// Execute service functions for 1st reflections.
	((LPA3DREFLECTIONS)pA3dReflections)->Service();

	// Terminate this thread.
	ExitThread(NO_ERROR);

	return NO_ERROR;
}


//===========================================================================
//
// IA3dReflections::CreateReflection
//
// Purpose: Create resources for reflection.
//
// Parameters:
//  dwNumRef        DWORD notification events count.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dReflections::CreateReflection(DWORD dwNumRef)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dReflections::CreateReflection(%u)"), dwNumRef);
	_ASSERTE(dwNumRef < A3D_MAX_SOURCE_REFLECTIONS);
	_ASSERTE(m_pDS);
	_ASSERTE(m_pDSB);
	_ASSERTE(!m_pRefsDSB[dwNumRef]);
#endif
	// Duplicate source sound buffer for reflection.
	HRESULT hr = m_pDS->DuplicateSoundBuffer(m_pDSB, &m_pRefsDSB[dwNumRef]);
	if (FAILED(hr))
		return hr;

	// Get DirectSound3DBuffer object for reflection.
	hr = m_pDSB->QueryInterface(IID_IDirectSound3DBuffer, (LPVOID *)&m_pRefsDS3DB[dwNumRef]);
	if (FAILED(hr))
	{
		// Release reflection sound buffer.
		m_pRefsDSB[dwNumRef]->Release();
		m_pRefsDSB[dwNumRef] = NULL;

		return hr;
	}

	return S_OK;
}


//===========================================================================
//
// IA3dReflections::SchedulePlay
//
// Purpose: Schedule play reflection sound buffers.
//
// Parameters:
//  dwNumRef        DWORD notification events count.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dReflections::SchedulePlay(DWORD dwNotifyCount)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dReflections::SchedulePlay(%u)"), dwNotifyCount);
	_ASSERTE(dwNotifyCount <= A3D_MAX_SOURCE_REFLECTIONS);
	_ASSERTE(m_pDSN);
#endif
	// Set notifications for source sound buffer.
	HRESULT hr = m_pDSN->SetNotificationPositions(dwNotifyCount, m_DSBPN);
#ifdef _DEBUG
	LogMsg(TEXT("...SetNotificationPositions(%u)=%s"), dwNotifyCount, Result(hr));
#endif
	if (FAILED(hr))
		return hr;

	// Exist reflection service thread handle.
	if (m_hThread)
	{
		DWORD dwExitCode;

		// Check executing service thread.
		if (GetExitCodeThread(m_hThread, &dwExitCode) && STILL_ACTIVE == dwExitCode)
		{
			// Signal about change notification events set.
			if (!PulseEvent(m_DSBPN[0].hEventNotify))
				return E_FAIL;

			return S_OK;
		}
#ifdef _DEBUG
		LogMsg(TEXT("...GetExitCodeThread(%#x)=%#x"), m_hThread, dwExitCode);
#endif

		// Close previous thread handle.
		CloseHandle(m_hThread);
	}

	DWORD dwThreadId;

	// Create reflection service thread.
	m_hThread = CreateThread(NULL, 0, ServiceThread, this, 0, &dwThreadId);
#ifdef _DEBUG
	LogMsg(TEXT("...CreateThread(%#x)=%#x"), this, m_hThread);
#endif
	if (!m_hThread)
		return E_FAIL;

	// Set reflection service thread priority.
	if (!SetThreadPriority(m_hThread, THREAD_PRIORITY_TIME_CRITICAL))
		return E_FAIL;
#ifdef _DEBUG
	LogMsg(TEXT("...GetThreadPriority(%#x)=%#x"), m_hThread, GetThreadPriority(m_hThread));
#endif

	return S_OK;
}


//===========================================================================
//
// IA3dReflections::PlayWithLag
//
// Purpose: Start playing reflection sound buffer.
//
// Parameters:
//  dwNumRef        DWORD reflection number.
//  dwFlags         DWORD flags for playing.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dReflections::PlayWithLag(DWORD dwNumRef, DWORD dwFlags)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dReflections::PlayWithLag(%u,%#x)"), dwNumRef, dwFlags);
	_ASSERTE(dwNumRef < A3D_MAX_SOURCE_REFLECTIONS);
	_ASSERTE(dwFlags);
	_ASSERTE(m_pDSB);
	_ASSERTE(m_pRefsDSB[dwNumRef]);
#endif
	DWORD dwSourcePosition;

	// Get play position for source sound buffer.
	HRESULT hr = m_pDSB->GetCurrentPosition(&dwSourcePosition, NULL);
	if (FAILED(hr))
		return hr;
#ifdef _DEBUG
	LogMsg(TEXT("...SourcePosition=%u"), dwSourcePosition);
#endif

	DWORD dwRefStatus;

	// Get current status for reflection sound buffer.
	hr = m_pRefsDSB[dwNumRef]->GetStatus(&dwRefStatus);
	if (FAILED(hr))
		return hr;
#ifdef _DEBUG
	LogMsg(TEXT("...Status[%u]=%#x"), dwNumRef, dwRefStatus);
#endif

	// Reflection sound buffer playing now.
	if (dwRefStatus & DSBSTATUS_PLAYING)
	{
		DWORD dwRefPosition;

		// Get play position for reflection sound buffer.
		hr = m_pRefsDSB[dwNumRef]->GetCurrentPosition(&dwRefPosition, NULL);
		if (FAILED(hr))
			return hr;

		DWORD dwOffset, dwFrequency;

		// Calculate offset between source and reflection play positions.
		if (dwSourcePosition >= dwRefPosition)
			dwOffset = dwSourcePosition - dwRefPosition;
		else
			dwOffset = m_dwBufferSize - dwRefPosition + dwSourcePosition;
#ifdef _DEBUG
		LogMsg(TEXT("...Offset[%u]=%u(%u)"), dwNumRef,
			dwOffset, m_DSBPN[dwNumRef + 1].dwOffset);
#endif

		// Calculate maximum offset between source and reflection positions.
		DWORD dwMaxDelta = (m_dwSourceFrequency * m_dwBytesPerSample *
			A3DREF_DELAY_PRECISION) / 1024;
#ifdef _DEBUG
		LogMsg(TEXT("...MaxDelta=%u"), dwMaxDelta);
#endif

		// Calculate new frequency for reflection sound buffer.
		if (dwOffset > (m_DSBPN[dwNumRef + 1].dwOffset + dwMaxDelta))
			dwFrequency = (m_dwSourceFrequency * (128 + A3DREF_CHANGE_FREQUENCY)) / 128;
		else if (dwOffset < (m_DSBPN[dwNumRef + 1].dwOffset - dwMaxDelta))
			dwFrequency = (m_dwSourceFrequency * (128 - A3DREF_CHANGE_FREQUENCY)) / 128;
		else
			dwFrequency = m_dwSourceFrequency;

		// Set reflection sound buffer frequency.
		hr = m_pRefsDSB[dwNumRef]->SetFrequency(dwFrequency);
		if (FAILED(hr))
			return hr;
#ifdef _DEBUG
		LogMsg(TEXT("...Frequency[%u]=%u(%u)"), dwNumRef, dwFrequency, m_dwSourceFrequency);
#endif
	}
	else
	{
		// Set reflection sound buffer frequency.
		hr = m_pRefsDSB[dwNumRef]->SetFrequency(m_dwSourceFrequency);
		if (FAILED(hr))
			return hr;

		DWORD dwRefPosition;

		// Calculate play position for reflection sound buffer.
		if (dwSourcePosition >= m_DSBPN[dwNumRef + 1].dwOffset)
			dwRefPosition = dwSourcePosition - m_DSBPN[dwNumRef + 1].dwOffset;
		else
			dwRefPosition = m_dwBufferSize - m_DSBPN[dwNumRef + 1].dwOffset + dwSourcePosition;
#ifdef _DEBUG
		LogMsg(TEXT("...Offset[%u]=%u"), dwNumRef, m_DSBPN[dwNumRef + 1].dwOffset);
#endif

		// Set play position for reflection sound buffer.
		hr = m_pRefsDSB[dwNumRef]->SetCurrentPosition(dwRefPosition);
		if (FAILED(hr))
			return hr;
#ifdef _DEBUG
		LogMsg(TEXT("...Position[%u]=%u"), dwNumRef, dwRefPosition);
#endif
	}

	// Start playing reflection sound buffer.
	if (!(dwRefStatus & DSBSTATUS_PLAYING) ||
	(dwFlags & DSBSTATUS_LOOPING) != (dwRefStatus & DSBSTATUS_LOOPING))
		hr = m_pRefsDSB[dwNumRef]->Play(0, 0,
			(dwFlags & DSBSTATUS_LOOPING) ? DSBPLAY_LOOPING : 0);

	return hr;
}


//===========================================================================
//
// IA3dReflections::Service
//
// Purpose: Wait and start playing delayed reflections.
//
//===========================================================================
STDMETHODIMP_(VOID) IA3dReflections::Service()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dReflections::Service()"));
	_ASSERTE(m_pDSB);
#endif
	HANDLE hEventsList[A3D_MAX_SOURCE_REFLECTIONS + 1];

	// Save to set stop source sound buffer event.
	hEventsList[0] = m_DSBPN[0].hEventNotify;

	// Request reflections resources.
	EnterCriticalSection(&m_CS); 

	// Loop waiting notification events.
	for (;;)
	{
		// First exist only one event.
		DWORD dwEventCount = 1;

		DWORD dwRefList[A3D_MAX_SOURCE_REFLECTIONS];

		// Fill list events for all waiting reflections.
		for (UINT i = 0; i < A3D_MAX_SOURCE_REFLECTIONS; i++)
			if (m_DSBPN[i + 1].hEventNotify)
			{
				hEventsList[dwEventCount] = m_DSBPN[i + 1].hEventNotify;
				dwRefList[dwEventCount - 1] = i;
				dwEventCount++;
			}

		// Release reflections resources.
		LeaveCriticalSection(&m_CS);

#ifdef _DEBUG
		LogMsg(TEXT("...WaitForMultipleObjects(%u)..."), dwEventCount);
#endif
		// Wait any notification events for source sound buffer.
		DWORD dwNumObject = WaitForMultipleObjects(dwEventCount, hEventsList, FALSE, INFINITE);
#ifdef _DEBUG
		LogMsg(TEXT("...WaitForMultipleObjects(%u)=%u"), dwEventCount, dwNumObject);
#endif

		// Request reflections resources.
		EnterCriticalSection(&m_CS); 

		// Check correct returned number object.
		if (WAIT_OBJECT_0 > dwNumObject || (WAIT_OBJECT_0 + dwEventCount) < dwNumObject)
			break;

		// Event for stop source sound buffer signaled.
		if (WAIT_OBJECT_0 == dwNumObject)
		{
			// Check current stop source sound buffer event status.
			if (WaitForSingleObject(m_DSBPN[0].hEventNotify, 0) != WAIT_OBJECT_0)
				continue;

			// Reset stop source sound buffer event for future usage.
			ResetEvent(m_DSBPN[0].hEventNotify);
			break;
		}

		// Get signaled reflection number.
		DWORD dwNumRef = dwRefList[dwNumObject - WAIT_OBJECT_0 - 1];
#ifdef _DEBUG
		LogMsg(TEXT("...NumRef=%u"), dwNumRef);
#endif

		// Exist reflection sound buffer.
		if (m_pRefsDSB[dwNumRef])
		{
			DWORD dwSourceStatus;

			// Get current status for source sound buffer.
			if (FAILED(m_pDSB->GetStatus(&dwSourceStatus)))
				break;

			// Source sound buffer now stopped.
			if (!(dwSourceStatus & DSBSTATUS_PLAYING))
				break;

			// Play with lag reflection.
			if (FAILED(PlayWithLag(dwNumRef, dwSourceStatus)))
				break;
		}

		// Clear notification handle.
		CloseHandle(m_DSBPN[dwNumRef + 1].hEventNotify);
		m_DSBPN[dwNumRef + 1].hEventNotify = NULL;
	}

	// Stop all reflections.
	Stop();

	// Release reflections resources.
	LeaveCriticalSection(&m_CS);
}


//===========================================================================
//
// IA3dReflections::Reset
//
// Purpose: Release resources for reflection.
//
// Parameters:
//  dwNumRef        DWORD notification events count.
//
//===========================================================================
STDMETHODIMP_(VOID) IA3dReflections::Reset(DWORD dwNumRef)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dReflections::Reset(%u)"), dwNumRef);
	_ASSERTE(dwNumRef < A3D_MAX_SOURCE_REFLECTIONS);
	_ASSERTE(m_pRefsDSB[dwNumRef]);
	_ASSERTE(m_pRefsDS3DB[dwNumRef]);
#endif
	// Release DirectSound3DBuffer object for reflection.
	m_pRefsDS3DB[dwNumRef]->Release();

	// Stop and release reflection sound buffer.
	m_pRefsDSB[dwNumRef]->Stop();
	m_pRefsDSB[dwNumRef]->Release();
	m_pRefsDSB[dwNumRef] = NULL;

	// Clear notification handle for reflection.
	if (m_DSBPN[dwNumRef + 1].hEventNotify)
	{
		CloseHandle(m_DSBPN[dwNumRef + 1].hEventNotify);
		m_DSBPN[dwNumRef + 1].hEventNotify = NULL;
	}
}


//===========================================================================
//
// IA3dReflections::Stop
//
// Purpose: Release all resources for all reflections.
//
//===========================================================================
STDMETHODIMP_(VOID) IA3dReflections::Stop()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dReflections::Stop()"));
#endif
	// Reset all existing reflections.
	for (UINT i = 0; i < A3D_MAX_SOURCE_REFLECTIONS; i++)
		if (m_pRefsDSB[i])
			Reset(i);
}


//===========================================================================
//
// IA3dReflections::IA3dReflections
// IA3dReflections::~IA3dReflections
//
// Constructor Parameters:
//  None
//
//===========================================================================
IA3dReflections::IA3dReflections() :
	m_dwBufferSize(DSBSIZE_MAX),
	m_dwBytesPerSample(2),
	m_dwSourceFrequency(A3D_SAMPLE_RATE_1),
	m_pDS(NULL),
	m_pDSB(NULL),
	m_pDSN(NULL),
	m_hThread(NULL)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dReflections::IA3dReflections()=%u"), g_cObj + 1);
#endif
	// Zero big object members.
	ZeroMemory(m_pRefsDSB, sizeof(m_pRefsDSB));
	ZeroMemory(m_DSBPN, sizeof(m_DSBPN));

	// Initialize resources critical section.
	InitializeCriticalSection(&m_CS);

	// Increase object counter.
	InterlockedIncrement(&g_cObj);
}

IA3dReflections::~IA3dReflections()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dReflections::~IA3dReflections()=%u"), g_cObj - 1);
	_ASSERTE(g_cObj > 0);
#endif
	// Exit thread for reflection service.
	if (m_hThread)
	{
		DWORD dwExitCode;

		// Check executing service thread.
		if (GetExitCodeThread(m_hThread, &dwExitCode) && STILL_ACTIVE == dwExitCode)
		{
			// Signal about stop source sound buffer.
			SetEvent(m_DSBPN[0].hEventNotify);

			// Wait termination of thread and to kill it.
			if (WaitForSingleObject(m_hThread, A3DREF_MAX_WAIT_THREAD) != WAIT_OBJECT_0)
				TerminateThread(m_hThread, E_FAIL);
		}
#ifdef _DEBUG
		LogMsg(TEXT("...GetExitCodeThread(%#x)=%#x"), m_hThread, dwExitCode);
#endif

		// Close thread handle.
		CloseHandle(m_hThread);
	}

	// Stop all reflections.
	Stop();

	// Close stop source sound buffer event.
	if (m_DSBPN[0].hEventNotify)
		CloseHandle(m_DSBPN[0].hEventNotify);

	// Release DirectSoundNotify object.
	if (m_pDSN)
		m_pDSN->Release();

	// Delete resources critical section.
	DeleteCriticalSection(&m_CS);

	// Decrease object counter.
	InterlockedDecrement(&g_cObj);
}


//===========================================================================
//
// IA3dReflections::Initialize
//
// Purpose: Initialize necessary resources reflections.
//
// Parameters:
//  pDS             LPDIRECTSOUND to the parent object.
//  pDSB            LPDIRECTSOUNDBUFFER to the source sound buffer.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dReflections::Initialize(LPDIRECTSOUND pDS, LPDIRECTSOUNDBUFFER pDSB)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dReflections::Initialize(%#x,%#x)"), pDS, pDSB);
	_ASSERTE(pDS);
	_ASSERTE(pDSB);
#endif
	// Save parent DirectSound object and source sound buffer.
	m_pDS = pDS;
	m_pDSB = pDSB;

	DSCAPS DSCaps;
	DSCaps.dwSize = sizeof(DSCaps);

	// Get DirectSound capabilities.
	HRESULT hr = m_pDS->GetCaps(&DSCaps);
	if (FAILED(hr))
		return hr;

	// Check audio driver compatibility.
	if (!DSCaps.dwMaxHw3DAllBuffers || !(DSCaps.dwFlags & DSCAPS_CONTINUOUSRATE))
		return E_FAIL;

	DSBCAPS DSBCaps;
	DSBCaps.dwSize = sizeof(DSBCaps);

	// Get source sound buffer capabilities.
	hr = m_pDSB->GetCaps(&DSBCaps);
	if (FAILED(hr))
		return hr;
#ifdef _DEBUG
	LogMsg(TEXT("...Flags=%#x"), DSBCaps.dwFlags);
	if (DSBCaps.dwFlags & DSBCAPS_CTRL3D)
		LogMsg(TEXT("...DSBCAPS_CTRL3D"));
	if (DSBCaps.dwFlags & DSBCAPS_LOCHARDWARE)
		LogMsg(TEXT("...DSBCAPS_LOCHARDWARE"));
	if (DSBCaps.dwFlags & DSBCAPS_CTRLPOSITIONNOTIFY)
		LogMsg(TEXT("...DSBCAPS_CTRLPOSITIONNOTIFY"));
#endif

	// Check source sound buffer compatibility.
	if (!(DSBCaps.dwFlags & DSBCAPS_CTRL3D) ||
	!(DSBCaps.dwFlags & DSBCAPS_LOCHARDWARE) ||
	!(DSBCaps.dwFlags & DSBCAPS_CTRLPOSITIONNOTIFY))
		return E_FAIL;

	// Get size for source sound buffer.
	m_dwBufferSize = DSBCaps.dwBufferBytes;

	// Get DirectSoundNotify object.
	hr = m_pDSB->QueryInterface(IID_IDirectSoundNotify, (LPVOID *)&m_pDSN);
#ifdef _DEBUG
	LogMsg(TEXT("...QueryInterface(IID_IDirectSoundNotify)=%s"), Result(hr));
#endif
	if (FAILED(hr))
		return hr;

	// Create stop source sound buffer event.
	m_DSBPN[0].hEventNotify = CreateEvent(NULL, TRUE, FALSE, NULL);
	m_DSBPN[0].dwOffset = DSBPN_OFFSETSTOP;

	// Set one notification position.
	hr = m_pDSN->SetNotificationPositions(1, m_DSBPN);
#ifdef _DEBUG
	LogMsg(TEXT("...SetNotificationPositions(1)=%s"), Result(hr));
#endif
	if (FAILED(hr))
		return hr;

	// Prepare 3D position buffer for reflections.
	ZeroMemory(&m_DS3DBuffer, sizeof(m_DS3DBuffer));
	m_DS3DBuffer.dwSize = sizeof(m_DS3DBuffer);
	m_DS3DBuffer.dwInsideConeAngle = DS3D_DEFAULTCONEANGLE;
	m_DS3DBuffer.dwOutsideConeAngle = DS3D_DEFAULTCONEANGLE;
	m_DS3DBuffer.vConeOrientation.z = 1.0f;
	m_DS3DBuffer.flMinDistance = 0.5f;
	m_DS3DBuffer.flMaxDistance = 2.0f;
	m_DS3DBuffer.dwMode = DS3DMODE_HEADRELATIVE;

	return S_OK;
}


//===========================================================================
//
// IA3dReflections::SetA3dSuperCtrl
//
// Purpose: Set reflections data from control packet.
//
// Parameters:
//  pA3dCtrlSuper   LPA3DCTRL_SRC_SUPER pointer to control packet.
//  dwSourceFrequency DWORD frequency source sound buffer.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dReflections::SetA3dSuperCtrl(LPA3DCTRL_SRC_SUPER pA3dCtrlSuper,
	DWORD dwSourceFrequency)
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dReflections::SetA3dSuperCtrl(%#x,%u)..."),
		pA3dCtrlSuper, dwSourceFrequency);
	_ASSERTE(pA3dCtrlSuper && !IsBadReadPtr(pA3dCtrlSuper, sizeof(*pA3dCtrlSuper)));
	_ASSERTE(m_pDSB);
#endif
	DWORD dwSourceStatus;

	// Get current status for source sound buffer.
	HRESULT hr = m_pDSB->GetStatus(&dwSourceStatus);
	if (FAILED(hr))
		return hr;
#ifdef _DEBUG
	LogMsg(TEXT("...GetStatus()=%#x"), dwSourceStatus);

	if (dwSourceStatus & DSBSTATUS_PLAYING)
		LogMsg(TEXT("......DSBSTATUS_PLAYING"));
	if (dwSourceStatus & DSBSTATUS_BUFFERLOST)
		LogMsg(TEXT("......DSBSTATUS_BUFFERLOST"));
	if (dwSourceStatus & DSBSTATUS_LOOPING)
		LogMsg(TEXT("......DSBSTATUS_LOOPING"));
#endif

	WAVEFORMATEX wfxFormat;

	// Get source sound buffer format.
	hr = m_pDSB->GetFormat(&wfxFormat, sizeof(wfxFormat), NULL);
	if (FAILED(hr))
		return hr;

	// Request reflections resources.
	EnterCriticalSection(&m_CS);

	// Save source sound buffer frequency.
	m_dwSourceFrequency = dwSourceFrequency;

	// Save source sound buffer bytes per sample.
	m_dwBytesPerSample = wfxFormat.wBitsPerSample / 8;

	// Calculate half play speed in bytes for source sound buffer.
	A3DVAL fHalfSpeed = (m_dwSourceFrequency * m_dwBytesPerSample) * 0.5f;

	// First exist only notification event for stop.
	DWORD dwCounter = 1;

	// Enumerates all reflections.
	for (UINT i = 0; i < A3D_MAX_SOURCE_REFLECTIONS; i++)
	{
#ifdef _DEBUG
		// Reflection not tracked.
		if (pA3dCtrlSuper->Reflections[i].bEnable && pA3dCtrlSuper->Reflections[i].bMute)
			LogMsg(TEXT("...[%u].bMute=%u!"), i,
				pA3dCtrlSuper->Reflections[i].bMute);
#endif
		// Reflection not enable or not available now.
		if (!pA3dCtrlSuper->Reflections[i].bEnable ||
		!pA3dCtrlSuper->Reflections[i].bAvailable || pA3dCtrlSuper->Reflections[i].bMute)
		{
			// Reset reflection sound buffer.
			if (m_pRefsDSB[i])
				Reset(i);

			continue;
		}

		// Create reflection sound buffer.
		if (!m_pRefsDSB[i])
		{
			hr = CreateReflection(i);
			if (FAILED(hr))
				break;
		}

		// Calculate average sound buffer position.
		A3DVAL fAzim = (pA3dCtrlSuper->Reflections[i].LeftEar.fAzim +
			pA3dCtrlSuper->Reflections[i].RightEar.fAzim) * 0.5f;
		A3DVAL fElev = (pA3dCtrlSuper->Reflections[i].LeftEar.fElev +
			pA3dCtrlSuper->Reflections[i].RightEar.fElev) * 0.5f;

		// Calculate DirectSound3D sound buffer position.
		m_DS3DBuffer.vPosition.x = (D3DVALUE)-sin(fAzim);
		m_DS3DBuffer.vPosition.y = (D3DVALUE)sin(fElev);
		m_DS3DBuffer.vPosition.z = (D3DVALUE)cos(fAzim);
#ifdef _DEBUG
		LogMsg(TEXT("...[%u].vPosition.x=%g"), i, m_DS3DBuffer.vPosition.x);
		LogMsg(TEXT("...[%u].vPosition.y=%g"), i, m_DS3DBuffer.vPosition.y);
		LogMsg(TEXT("...[%u].vPosition.z=%g"), i, m_DS3DBuffer.vPosition.z);
#endif

		// Set reflection sound buffer position.
		hr = m_pRefsDS3DB[i]->SetAllParameters(&m_DS3DBuffer, DS3D_IMMEDIATE);
#ifdef _DEBUG
		LogMsg(TEXT("...[%u].SetAllParameters()=%s"), i, Result(hr));
#endif
		if (FAILED(hr))
			return hr;

		// Calculate average gain for reflection sound buffer.
		A3DVAL fGain = (pA3dCtrlSuper->Reflections[i].LeftEar.fGain +
			pA3dCtrlSuper->Reflections[i].RightEar.fGain) * 0.5f;

		// Correct gain for equalization effect.
		if (pA3dCtrlSuper->Reflections[i].fAlpha > 0.4f)
			fGain *= 1.67f - 1.67f * pA3dCtrlSuper->Reflections[i].fAlpha;

		// Calculate volume level for reflection sound buffer.
		LONG lVolume;
		if (fGain < 0.00001f)
			lVolume = DSBVOLUME_MIN;
		else
			lVolume = (LONG)(log10(fGain) * 2000.0f);

		// Set reflection sound buffer volume level.
		hr = m_pRefsDSB[i]->SetVolume(lVolume);
#ifdef _DEBUG
		LogMsg(TEXT("...[%u].SetVolume(%d)=%s"), i, lVolume, Result(hr));
#endif
		if (FAILED(hr))
			break;

		// Calculate average offset in sound buffer.
		DWORD dwOffset = (DWORD)((pA3dCtrlSuper->Reflections[i].LeftEar.fDelay +
			pA3dCtrlSuper->Reflections[i].RightEar.fDelay) * fHalfSpeed);
#ifdef _DEBUG
		LogMsg(TEXT("...[%u].fDelay=%g"), i, 
			(pA3dCtrlSuper->Reflections[i].LeftEar.fDelay +
			pA3dCtrlSuper->Reflections[i].RightEar.fDelay) * 0.5f);
#endif

		// Align offset value.
		if (dwOffset % m_dwBytesPerSample)
			dwOffset += m_dwBytesPerSample - dwOffset % m_dwBytesPerSample;

		// Loop for correction offset.
		bool bOffsetChange;
		do {
			bOffsetChange = false;

			// Comparison with the previous offsets.
			for (UINT j = 0; j < i; j++)
				if (m_DSBPN[j + 1].dwOffset == dwOffset)
				{
					dwOffset += m_dwBytesPerSample;
					bOffsetChange = true;
					break;
				}

			// Correction offset for overflow sound buffer size.
			if (dwOffset >= m_dwBufferSize)
				dwOffset %= m_dwBufferSize;
		} while (bOffsetChange);

		// Save true offset in structure notification positions.
		m_DSBPN[i + 1].dwOffset = dwOffset;
#ifdef _DEBUG
		LogMsg(TEXT("...[%u].dwOffset=%u"), i, dwOffset);
#endif

		// Source sound buffer now playing.
		if (dwSourceStatus & DSBSTATUS_PLAYING)
		{
			// Play reflection not scheduled in future.
			if (!m_DSBPN[i + 1].hEventNotify)
			{
				// Play with lag reflection.
				hr = PlayWithLag(i, dwSourceStatus);
				if (FAILED(hr))
					break;
			}
		}
		else
		{
			// Create event for notification.
			if (!m_DSBPN[i + 1].hEventNotify)
				m_DSBPN[i + 1].hEventNotify = CreateEvent(NULL, TRUE, FALSE, NULL);

			// Set new number of notifications.
			dwCounter = i + 2;
		}
	}

	// Source sound buffer now not playing.
	if (SUCCEEDED(hr) && !(dwSourceStatus & DSBSTATUS_PLAYING))
		hr = SchedulePlay(dwCounter);

	// For failed return stop all reflections.
	if (FAILED(hr))
		Stop();

	// Release reflections resources.
	LeaveCriticalSection(&m_CS);

	return hr;
}


//===========================================================================
//
// IA3dReflections::TrackDelay
//
// Purpose: Track delay for reflection sound buffers.
//
// Return: S_OK if successful, error otherwise.
//
//===========================================================================
STDMETHODIMP IA3dReflections::TrackDelay()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dReflections::TrackDelay()..."));
	_ASSERTE(m_pDSB);
#endif
	DWORD dwSourceStatus;

	// Get current status for source sound buffer.
	HRESULT hr = m_pDSB->GetStatus(&dwSourceStatus);
	if (FAILED(hr))
		return hr;
#ifdef _DEBUG
	LogMsg(TEXT("...GetStatus()=%#x"), dwSourceStatus);

	if (dwSourceStatus & DSBSTATUS_PLAYING)
		LogMsg(TEXT("......DSBSTATUS_PLAYING"));
	if (dwSourceStatus & DSBSTATUS_BUFFERLOST)
		LogMsg(TEXT("......DSBSTATUS_BUFFERLOST"));
	if (dwSourceStatus & DSBSTATUS_LOOPING)
		LogMsg(TEXT("......DSBSTATUS_LOOPING"));
#endif

	// Source sound buffer now not playing.
	if (!(dwSourceStatus & DSBSTATUS_PLAYING))
		return S_OK;

	// Request reflections resources.
	EnterCriticalSection(&m_CS);

	// Enumerates all reflections.
	for (UINT i = 0; i < A3D_MAX_SOURCE_REFLECTIONS; i++)
	{
		// Reflection exist and not scheduled play in future.
		if (m_pRefsDSB[i] && !m_DSBPN[i + 1].hEventNotify)
		{
			// Play with lag reflection.
			hr = PlayWithLag(i, dwSourceStatus);
			if (FAILED(hr))
				break;
		}
	}

	// For failed return stop all reflections.
	if (FAILED(hr))
		Stop();

	// Release reflections resources.
	LeaveCriticalSection(&m_CS);

	return hr;
}


//===========================================================================
//
// IA3dReflections::ReadyForService
//
// Purpose: Get reflections count.
//
// Return: Number current ready reflections.
//
//===========================================================================
STDMETHODIMP_(DWORD) IA3dReflections::ReadyForService()
{
#ifdef _DEBUG
	LogMsg(TEXT("IA3dReflections::ReadyForService()..."));
#endif
	DWORD dwCounter = 0;

	// Request reflections resources.
	EnterCriticalSection(&m_CS);

	// Enumerates all reflections.
	for (UINT i = 0; i < A3D_MAX_SOURCE_REFLECTIONS; i++)
		if (m_pRefsDSB[i])
			dwCounter++;

	// Release reflections resources.
	LeaveCriticalSection(&m_CS);

#ifdef _DEBUG
	LogMsg(TEXT("...ReadyForService()=%u"), dwCounter);
#endif
	return dwCounter;
}
