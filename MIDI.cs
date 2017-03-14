/*
    Permission is granted to copy, distribute and/or modify this document
    under the terms of the GNU Free Documentation License, Version 1.3
    or any later version published by the Free Software Foundation;
    with no Invariant Sections, no Front-Cover Texts, and no Back-Cover Texts.
    A copy of the license is included in the section entitled "GNU
    Free Documentation License".
*/


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;


namespace MIDIWrapper
{

    

    //This module containes all midi API functions
    static class MIDI
    {
        public const short MAXPNAMELEN = 32; //  max product name length (including NULL)
        // General error return values
        public const short MMSYSERR_BASE = 0;
        //  no error
        public const short MMSYSERR_NOERROR = 0;
        //  unspecified error
        public const int MMSYSERR_ERROR = (MMSYSERR_BASE + 1);
        //  device ID out of range
        public const int MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2);
        //  driver failed enable
        public const int MMSYSERR_NOTENABLED = (MMSYSERR_BASE + 3);
        //  device already allocated
        public const int MMSYSERR_ALLOCATED = (MMSYSERR_BASE + 4);
        //  device handle is invalid
        public const int MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5);
        //  no device driver present
        public const int MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6);
        //  memory allocation error
        public const int MMSYSERR_NOMEM = (MMSYSERR_BASE + 7);
        //  function isn't supported
        public const int MMSYSERR_NOTSUPPORTED = (MMSYSERR_BASE + 8);
        //  error value out of range
        public const int MMSYSERR_BADERRNUM = (MMSYSERR_BASE + 9);
        //  invalid flag passed
        public const int MMSYSERR_INVALFLAG = (MMSYSERR_BASE + 10);
        //  invalid parameter passed
        public const int MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11);
        //  handle being used
        public const int MMSYSERR_HANDLEBUSY = (MMSYSERR_BASE + 12);

        //  simultaneously on another
        //  thread (eg callback)
        //  "Specified alias not found in WIN.INI
        public const int MMSYSERR_INVALIDALIAS = (MMSYSERR_BASE + 13);
        //  last error in range
        public const int MMSYSERR_LASTERROR = (MMSYSERR_BASE + 13);

        //  flags for dwFlags field of MIDIHDR structure
        //  done bit
        public const short MHDR_DONE = 0x1;
        //  set if header prepared
        public const short MHDR_PREPARED = 0x2;
        //  reserved for driver
        public const short MHDR_INQUEUE = 0x4;
        //  valid flags / ;Internal /
        public const short MHDR_VALID = 0x7;

        //  flags used with waveOutOpen(), waveInOpen(), midiInOpen(), and
        //  midiOutOpen() to specify the type of the dwCallback parameter.
        //  callback type mask
        public const int CALLBACK_TYPEMASK = 0x70000;
        //  no callback
        public const short CALLBACK_NULL = 0x0;
        //  dwCallback is a HWND
        public const int CALLBACK_WINDOW = 0x10000;
        //  dwCallback is a HTASK
        public const int CALLBACK_TASK = 0x20000;
        //  dwCallback is a FARPROC
        public const int CALLBACK_FUNCTION = 0x30000;

        //  manufacturer IDs
        //  Microsoft Corp.
        public const short MM_MICROSOFT = 1;

        //  product IDs
        //  MIDI Mapper
        public const short MM_MIDI_MAPPER = 1;
        //  Wave Mapper
        public const short MM_WAVE_MAPPER = 2;

        //  Sound Blaster MIDI output port
        public const short MM_SNDBLST_MIDIOUT = 3;
        //  Sound Blaster MIDI input port
        public const short MM_SNDBLST_MIDIIN = 4;
        //  Sound Blaster internal synthesizer
        public const short MM_SNDBLST_SYNTH = 5;
        //  Sound Blaster waveform output
        public const short MM_SNDBLST_WAVEOUT = 6;
        //  Sound Blaster waveform input
        public const short MM_SNDBLST_WAVEIN = 7;

        //  Ad Lib-compatible synthesizer
        public const short MM_ADLIB = 9;

        //  MPU401-compatible MIDI output port
        public const short MM_MPU401_MIDIOUT = 10;
        //  MPU401-compatible MIDI input port
        public const short MM_MPU401_MIDIIN = 11;

        //  Joystick adapter
        public const short MM_PC_JOYSTICK = 12;

        public const short MIDI_IO_STATUS = 0x20;
        //  MIDI input
        public const short MM_MIM_OPEN = 0x3c1;
        public const short MM_MIM_CLOSE = 0x3c2;
        public const short MM_MIM_DATA = 0x3c3;
        public const short MM_MIM_LONGDATA = 0x3c4;
        public const short MM_MIM_ERROR = 0x3c5;
        public const short MM_MIM_LONGERROR = 0x3c6;

        public const short MM_MIM_MOREDATA = 0x3cc;
        //  MIDI output
        public const short MM_MOM_OPEN = 0x3c7;
        public const short MM_MOM_CLOSE = 0x3c8;

        public const short MM_MOM_DONE = 0x3c9;
        //----------------------------------------------------------------

        //MIDI Mapper

        public const short MIDI_MAPPER = -1;
        //  flags for wTechnology field of MIDIOUTCAPS structure
        //  output port
        public const byte MOD_MIDIPORT = 1;
        //  generic internal synth
        public const byte MOD_SYNTH = 2;
        //  square wave internal synth
        public const byte MOD_SQSYNTH = 3;
        //  FM internal synth
        public const byte MOD_FMSYNTH = 4;
        //  MIDI mapper
        public const byte MOD_MAPPER = 5;

        //  flags for dwSupport field of MIDIOUTCAPS
        //  supports volume control
        public const byte MIDICAPS_VOLUME = 0x1;
        //  separate left-right volume control
        public const byte MIDICAPS_LRVOLUME = 0x2;

        public const byte MIDICAPS_CACHE = 0x4;
        //' MIDI Controller Numbers Constants
        public const byte MOD_WHEEL = 1;
        public const byte BREATH_CONTROLLER = 2;
        public const byte FOOT_CONTROLLER = 4;
        public const byte PORTAMENTO_TIME = 5;
        public const byte MAIN_VOLUME = 7;
        public const byte BALANCE = 8;
        public const byte PAN = 10;
        public const byte EXPRESS_CONTROLLER = 11;
        public const byte DAMPER_PEDAL = 64;
        public const byte PORTAMENTO = 65;
        public const byte SOSTENUTO = 66;
        public const byte SOFT_PEDAL = 67;
        public const byte HOLD_2 = 69;
        public const byte EXTERNAL_FX_DEPTH = 91;
        public const byte TREMELO_DEPTH = 92;
        public const byte CHORUS_DEPTH = 93;
        public const byte DETUNE_DEPTH = 94;
        public const byte PHASER_DEPTH = 95;
        public const byte DATA_INCREMENT = 96;

        public const byte DATA_DECREMENT = 97;
        // MIDI status messages
        public const byte NOTE_OFF = 0x80;
        public const byte NOTE_ON = 0x90;
        public const byte POLY_KEY_PRESS = 0xa0;
        public const byte CONTROLLER_CHANGE = 0xb0;
        public const byte PROGRAM_CHANGE = 0xc0;
        public const byte CHANNEL_PRESSURE = 0xd0;
        public const byte PITCH_BEND = 0xe0;


        public struct MIDIEVENT
        {
            //  Ticks since last event
            public int dwDeltaTime;
            //  Reserved; must be zero
            public int dwStreamID;
            //  Event type and parameters
            public int dwEvent;
            [VBFixedArray(1)]
            //  Parameters if this is a long event
            public int[] dwParms;
            public void Initialize()
            {
                dwParms = new int[2];
            }
        }

        public struct MIDIHDR
        {
            public string lpData;
            public int dwBufferLength;
            public int dwBytesRecorded;
            public int dwUser;
            public int dwFlags;
            public int lpNext;
            public int Reserved;
            public int dwOffset;
            [VBFixedArray(4)]

            public int[] dwReserved;
            public void Initialize()
            {
                dwReserved = new int[5];
            }
        }

        public struct MIDIINCAPS
        {
            public short wMid;
            public short wPid;
            public int vDriverVersion;
            [VBFixedString(MAXPNAMELEN), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst = MAXPNAMELEN)]
            public string szPname;
        }

        public struct MIDIOUTCAPS
        {
            public short wMid;
            public short wPid;
            public int vDriverVersion;
            [VBFixedString(MAXPNAMELEN), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst = MAXPNAMELEN)]
            public string szPname;
            public short wTechnology;
            public short wVoices;
            public short wNotes;
            public short wChannelMask;
            public int dwSupport;
        }

        public struct MIDIPROPTEMPO
        {
            public int cbStruct;
            public int dwTempo;
        }

        public struct MIDIPROPTIMEDIV
        {
            public int cbStruct;
            public int dwTimeDiv;
        }

        public struct MIDISTRMBUFFVER
        {
            //  Stream buffer format version
            public int dwVersion;
            //  Manufacturer ID as defined in MMREG.H
            public int dwMid;
            //  Manufacturer version for custom ext
            public int dwOEMVersion;
        }




        

        // MIDI API Functions for Windows 95

        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiConnect(int hmi, int hmo, ref Int32 pReserved);

        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiDisconnect(int hmi, int hmo, ref Int32 pReserved);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiInAddBuffer(int hMidiIN, ref MIDIHDR lpMidiInHdr, int uSize);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiInClose(int hMidiIN);
        
        [DllImport("winmm.dll", EntryPoint = "midiInGetDevCapsA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiInGetDevCaps(int uDeviceID, ref MIDIINCAPS lpCaps, int uSize);
        
        [DllImport("winmm.dll", EntryPoint = "midiInGetErrorTextA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiInGetErrorText(int err_Renamed, string lpText, int uSize);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiInGetID(int hMidiIN, ref string DeviceName);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiInGetNumDevs();
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiInMessage(int hMidiIN, int Msg, int dw1, int dw2);
        
        public delegate void MidiDelegate(Int32 MidiInHandle, Int32 wMsg, Int32 Instance, Int32 wParam, Int32 lParam);


        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiInOpen(ref int lphMidiIn, int uDeviceID, [MarshalAs(UnmanagedType.FunctionPtr)]MidiDelegate dwCallback, int dwInstance, int dwFlags);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiInPrepareHeader(int hMidiIN, ref MIDIHDR lpMidiInHdr, int uSize);
 
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiInReset(int hMidiIN);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiInStart(int hMidiIN);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiInStop(int hMidiIN);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiInUnprepareHeader(int hMidiIN, ref MIDIHDR lpMidiInHdr, int uSize);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiOutCacheDrumPatches(int hMidiOUT, int uPatch, ref int lpKeyArray, int uFlags);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiOutCachePatches(int hMidiOUT, int uBank, ref int lpPatchArray, int uFlags);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiOutClose(int hMidiOUT);
        
        [DllImport("winmm.dll", EntryPoint = "midiOutGetDevCapsA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiOutGetDevCaps(int uDeviceID, ref MIDIOUTCAPS lpCaps, int uSize);
        
        [DllImport("winmm.dll", EntryPoint = "midiOutGetErrorTextA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiOutGetErrorText(int err_Renamed, string lpText, int uSize);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiOutGetID(int hMidiOUT, ref string DeviceName);
        
        [DllImport("winmm", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern short midiOutGetNumDevs();
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiOutGetVolume(int uDeviceID, ref int lpdwVolume);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiOutLongMsg(int hMidiOUT, ref MIDIHDR lpMidiOutHdr, int uSize);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiOutMessage(int hMidiOUT, int Msg, int dw1, int dw2);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiOutOpen(ref int lphMidiOut, int uDeviceID, int dwCallback, int dwInstance, int dwFlags);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiOutPrepareHeader(int hMidiOUT, ref MIDIHDR lpMidiOutHdr, int uSize);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiOutReset(int hMidiOUT);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiOutSetVolume(int uDeviceID, int dwVolume);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiOutShortMsg(int hMidiOUT, int dwMsg);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiOutUnprepareHeader(int hMidiOUT, ref MIDIHDR lpMidiOutHdr, int uSize);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiStreamClose(int hms);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiStreamOpen(ref int phms, ref int puDeviceID, int cMidi, int dwCallback, int dwInstance, int fdwOpen);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiStreamOut(int hms, ref MIDIHDR pmh, int cbmh);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiStreamPause(int hms);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiStreamProperty(int hms, ref byte lppropdata, int dwProperty);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiStreamRestart(int hms);
        
        [DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int midiStreamStop(int hms);

    }
}
