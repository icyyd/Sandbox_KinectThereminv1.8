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
using System.Collections;
using Microsoft.VisualBasic;
using System.Diagnostics;


namespace MIDIWrapper
{
    [Serializable()]
    public class Instrument : System.ComponentModel.Component
    {
        #region " Component Designer generated code "
        public Instrument(): base()
        {
            //This call is required by the Component Designer.
            InitializeComponent();
            //Add any initialization after the InitializeComponent() call

            dlgMIDIIn = new MIDI.MidiDelegate(MidiInProc);
        }

        //Component overrides dispose to clean up the component list.
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                this.Close();
                if ((components != null))
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        //Required by the Component Designer

        private System.ComponentModel.IContainer components;
        //NOTE: The following procedure is required by the Component Designer
        //It can be modified using the Component Designer.
        //Do not modify it using the code editor.
        [System.Diagnostics.DebuggerStepThrough()]

        private void InitializeComponent()
        {
        }

        #endregion

        #region " Private Declarations and Events "

        // constant definitions

        private const byte NOTE_ANZ  = 5;
        private byte[] whTonleiter = {0x00,0x03,0x05,0x07,0x0a};


        //-- Returns True if one or more MIDI ports are open

        private bool _Engaged = false;
        //-- Device numbers used internally
        private Int32 _OutputDeviceID = -1;

        private Int32 _InputDeviceID = -1;
        //-- MIDI Input and Output port handles returned by Windows
        private int hMidiOUT;

        private int hMidiIN;
        //-- This tool has an automatic note shutoff feature. When you set the 
        //   NoteDuration property, you can just call PlayNote and the notes will
        //   be cut off (NOTE_OFF sent) after NoteDuration number of seconds
        private ArrayList notesToTurnOff = new ArrayList();
        public struct NotesOff
        {
            public byte Note;
            public DateTime ShutoffTime;
        }

        //-- Receive event happens on all MIDI data when there is no output device open
        public event ReceiveEventHandler Receive;
        public delegate void ReceiveEventHandler(ref byte Channel, ref byte Status, ref byte Data1, ref byte Data2, ref byte Data3);
        //-- NoteOn event happens on Note On commands when there is no output device open
        public event NoteOnEventHandler NoteOn;
        public delegate void NoteOnEventHandler(ref byte Channel, ref byte Note, ref byte Velocity, ref bool Cancel);
        //-- NoteOff event happens on Note Off commands when there is no output device open
        public event NoteOffEventHandler NoteOff;
        public delegate void NoteOffEventHandler(ref byte Channel, ref byte Note, ref byte Velocity, ref bool Cancel);

        #endregion

        #region " Open, Close, and MIDI Exception "

        //-- This delegate is used for the MIDI callback function

        private MIDI.MidiDelegate dlgMIDIIn = null;
        //-- This is a private function to open and close a MIDI Input port. 
        private Int32 OpenMIDIInPort(ref int DeviceID, bool Open)
        {
            int midiError = 0;
            if (Open == true)
            {
                //-- This call opens the MIDI port using a callback function (MidiInProc)
                midiError = MIDI.midiInOpen(ref hMidiIN, DeviceID, dlgMIDIIn, 0, MIDI.CALLBACK_FUNCTION);
                if (midiError != MIDI.MMSYSERR_NOERROR)
                {
                    
                    ThrowMidiException("midiIN_Open", ref midiError);
                }
                else
                {
                    midiError = MIDI.midiInStart(hMidiIN);
                    if (midiError != MIDI.MMSYSERR_NOERROR)
                    {
                        ThrowMidiException("midiIN_Start", ref midiError);
                    }
                }
            }
            else
            {
                if (hMidiIN != 0)
                {
                    midiError = MIDI.midiInStop(hMidiIN);
                    if (midiError != MIDI.MMSYSERR_NOERROR)
                    {
                        ThrowMidiException("midiIN_Start", ref midiError);
                    }
                    else
                    {
                        midiError = MIDI.midiInClose(hMidiIN);
                        if (midiError != MIDI.MMSYSERR_NOERROR)
                        {
                            ThrowMidiException("midiIN_Close", ref midiError);
                        }
                        else
                        {
                            hMidiIN = 0;
                        }
                    }
                }
            }
            return hMidiIN;
        }

        //-- This function opens a MIDI Output port. No callback is needed
        private Int32 OpenMIDIOutPort(ref int DeviceID, bool Open)
        {
            int midiError = 0;
            if (Open == true)
            {
                //midiError = MIDI.midiOutOpen(ref hMidiOUT, DeviceID, VariantType.Null, 0, MIDI.CALLBACK_NULL);
                midiError = MIDI.midiOutOpen(ref hMidiOUT, DeviceID, (int)VariantType.Null, 0, MIDI.CALLBACK_NULL);
                if (midiError != MIDI.MMSYSERR_NOERROR)
                {
                    ThrowMidiException("midiOUT_Open", ref midiError);
                }
            }
            else
            {
                if (hMidiOUT != 0)
                {
                    midiError = MIDI.midiOutClose(hMidiOUT);
                    hMidiOUT = 0;
                    if (midiError != MIDI.MMSYSERR_NOERROR)
                    {
                        ThrowMidiException("midiOUT_Close", ref midiError);
                    }
                }
            }
            return hMidiOUT;
        }

        //-- Code to get the last MIDI message and throw as an exception
        private void ThrowMidiException(string InFunct, ref Int32 MMErr)
        {
            string Msg = Strings.Space(255);
            if (Strings.InStr(1, InFunct, "out", CompareMethod.Text) == 0)
            {
                MIDI.midiInGetErrorText(MMErr, Msg, 255);
            }
            else
            {
                MIDI.midiOutGetErrorText(MMErr, Msg, 255);
            }
            Msg = InFunct + Constants.vbCrLf + Msg + Constants.vbCrLf;
            switch (MMErr)
            {
                case MIDI.MMSYSERR_NOERROR:
                    Msg = Msg + "no error";
                    break;
                
                case MIDI.MMSYSERR_ERROR:  
                Msg = Msg + "unspecified error";
                    break;
                case MIDI.MMSYSERR_BADDEVICEID:
                    Msg = Msg + "device ID out of range";
                    break;
                case MIDI.MMSYSERR_NOTENABLED:
                    Msg = Msg + "driver failed enable";
                    break;
                case MIDI.MMSYSERR_ALLOCATED:
                    Msg = Msg + "device already allocated";
                    break;
                case MIDI.MMSYSERR_INVALHANDLE:
                    Msg = Msg + "device handle is invalid";
                    break;
                case MIDI.MMSYSERR_NODRIVER:
                    Msg = Msg + "no device driver present";
                    break;
                case MIDI.MMSYSERR_NOMEM:
                    Msg = Msg + "memory allocation error";
                    break;
                case MIDI.MMSYSERR_NOTSUPPORTED:
                    Msg = Msg + "function isn't supported";
                    break;
                case MIDI.MMSYSERR_BADERRNUM:
                    Msg = Msg + "error value out of range";
                    break;
                case MIDI.MMSYSERR_INVALFLAG:
                    Msg = Msg + "invalid flag passed";
                    break;
                case MIDI.MMSYSERR_INVALPARAM:
                    Msg = Msg + "invalid parameter passed";
                    break;
                case MIDI.MMSYSERR_HANDLEBUSY:
                    Msg = Msg + "handle being used simultaneously on another thread (eg callback)";
                    break;
                case MIDI.MMSYSERR_INVALIDALIAS:
                    Msg = Msg + "Specified alias not found in WIN.INI";
                    break;
                //case MIDI.MMSYSERR_LASTERROR:
                //    Msg = Msg + "last error in range";
                //    break;
            }
            throw new Exception(Msg);
        }

        #endregion

        #region " Callbacks and Finalizer "

        //-- This is the Input proc that gets called when MIDI Data is received from the
        //   input device.
        protected void MidiInProc(Int32 MidiInHandle, Int32 NewMsg, Int32 Instance, Int32 wParam, Int32 lParam)
        {
            byte chan = 0;
            byte Msg = 0;
            byte Status = 0;
            byte Data1 = 0;
            byte Data2 = 0;
            byte Data3 = 0;
     //       int MidiStatus = 0;
            bool Cancel = false;

            //-- We're only interested in MIDI Data messages
            if (NewMsg == MIDI.MM_MIM_DATA)
            {
                //-- Parse the data into a message byte and three data bytes
                SplitInt32(wParam, ref Msg, ref Data1, ref Data2, ref Data3);
                //-- The message byte is a combination of the channel and a Status byte. Parse
                SplitByte(Msg, ref chan, ref Status);

                Trace.WriteLine(" In: " + chan.ToString() + " " + Status.ToString() + " " + Data1.ToString() + " " + Data2.ToString() + " " + Data3.ToString());

                //-- Is this coming in on our Input Channel?

                if (chan == _InputChannel)
                {
                    //-- What MIDI Command was sent?
                    switch (Status)
                    {
                        case MIDI.NOTE_ON:
                        case MIDI.NOTE_OFF:
                            //-- Transpose the note
                            Int32 DTest = Convert.ToInt32(Data1) + Transpose;
                            if (DTest < 0)
                            {
                                DTest = 0;
                            }
                            else if (DTest > 127)
                            {
                                DTest = 127;
                            }
                            Data1 = GetByte(DTest);
                            //-- No output device?
                            if (hMidiOUT == 0)
                            {
                                //-- Fire the appropriate event
                                if (Status == MIDI.NOTE_ON)
                                {
                                    if (NoteOn != null)
                                    {
                                        NoteOn(ref chan, ref Data1, ref Data2, ref Cancel);
                                    }
                                }
                                else if (Status == MIDI.NOTE_OFF)
                                {
                                    if (NoteOff != null)
                                    {
                                        NoteOff(ref chan, ref Data1, ref Data2, ref Cancel);
                                    }
                                }
                            }
                            else
                            {
                                //-- Change the channel to the output channel
                                chan = _OutputChannel;
                            }
                            break;
                        case MIDI.CHANNEL_PRESSURE:
                            //-- Channel Pressure
                            if (FilterAfterTouch == true)
                            {
                                //-- Cancel this data
                                Cancel = true;
                            }
                            else if (hMidiOUT == 0)
                            {
                                //-- No output device. Fire the Receive event
                                if (Receive != null)
                                {
                                    Receive(ref chan, ref Status, ref Data1, ref Data2, ref Data3);
                                }
                            }
                            else
                            {
                                //-- Change the channel to the output channel
                                chan = _OutputChannel;
                            }
                            break;
                        default:
                            if (hMidiOUT == 0)
                            {
                                //-- No output device. Fire the Receive event
                                if (Receive != null)
                                {
                                    Receive(ref chan, ref Status, ref Data1, ref Data2, ref Data3);
                                }
                            }
                            else
                            {
                                //-- Change the channel to the output channel
                                chan = _OutputChannel;
                            }
                            break;
                    }
                }

                //-- Prepare the message 
                NewMsg = StuffByte(Status, chan);
                lParam = StuffInt32(Convert.ToByte(NewMsg), Data1, Data2, Data3);

                //-- The programmer can set Cancel to true to cancel this note.
                if (Cancel == false)
                {
                    if (hMidiOUT != 0)
                    {
                        //-- We have an output device!
                        Trace.WriteLine("Out: " + chan.ToString() + " " + Status.ToString() + " " + Data1.ToString() + " " + Data2.ToString() + " " + Data3.ToString() + " lParam=" + lParam.ToString());
                        //-- Send the MIDI data out the output device
                        MIDI.midiOutShortMsg(hMidiOUT, lParam);
                    }
                }
            }

        }

        //-- This is called by the thread to shut off a note after it has played.
        //   Only called when NoteDuration is set to non-zero
        protected void ShutoffNoteCallback()
        {
            NotesOff no = default(NotesOff);
            bool there = false;
            byte Msg = 0;
            Int32 MidiMsg = default(Int32);

            while (!(_Engaged == false))
            {
                lock (notesToTurnOff)
                {
                    if (notesToTurnOff.Count > 0)
                    {
                        no = (NotesOff)notesToTurnOff[0];
                        there = true;
                    }
                    else
                    {
                        there = false;
                    }
                }

                if (there)
                {
                    while (!(DateAndTime.Now >= no.ShutoffTime))
                    {
                        System.Threading.Thread.Sleep(1);
                        if (_Engaged == false)
                        {
                            notesToTurnOff.Clear();
                            ev.Set();
                            return;
                        }
                    }
                    Msg = StuffByte(GetByte(MIDIStatusMessages.NoteOff), _OutputChannel);
                    MidiMsg = StuffInt32(Msg, no.Note, 64, 0);
                    try
                    {
                        MIDI.midiOutShortMsg(hMidiOUT, MidiMsg);
                    }
                    catch (Exception )
                    {
                        break; // TODO: might not be correct. Was : Exit Do
                    }
                    lock (notesToTurnOff)
                    {
                        notesToTurnOff.RemoveAt(0);
                    }
                }
                System.Threading.Thread.Sleep(1);
            }
            notesToTurnOff.Clear();
            ev.Set();
        }

        ~Instrument()
        {
            MyNotesOff();
            //-- This is required so that the delegate is not garbage collected
            GC.KeepAlive(dlgMIDIIn);
            //base.Finalize();
        }

        #endregion

        #region " Numeric Conversion Utilities "
  
  
        public byte StuffByte(byte nib1, byte nib2)
        {
            //-- Stuffs two 4-bit values into a byte
            byte _16 = 16;
            return Convert.ToByte((nib1 * _16) + nib2);
        }

        public void SplitByte(byte OneByte, ref byte nib1, ref byte nib2)
        {
            //-- Splits a byte into two 4-bit values
            byte _15 = 15;
            byte _16 = 16;

            nib1 = Convert.ToByte(OneByte & _15);
            nib2 = Convert.ToByte(OneByte / _16);
        }

        public Int32 StuffInt32(byte B1, byte B2, byte B3, byte B4)
        {
            //-- Stuffs four bytes into an Int32
            int ret = 0;
            byte[] b = {
			B1,
			B2,
			B3,
			B4
		};
            ret = BitConverter.ToInt32(b, 0);
            return ret;
        }

        public void SplitInt32(int Data, ref byte B1, ref byte B2, ref byte B3, ref byte B4)
        {
            //-- Splits an Int32 into four bytes
            byte[] b = BitConverter.GetBytes(Data);
            B1 = b[0];
            B2 = b[1];
            B3 = b[2];
            B4 = b[3];
        }

        #endregion

        #region " Public Methods "

        //-- Return an InstrumentData object, containing only the 
        //   property settings of this instrument
        public InstrumentData GetSettings()
        {
            InstrumentData Data = new InstrumentData();
            var _with1 = Data;
            _with1.FilterAftertouch = this.FilterAfterTouch;
            _with1.InputChannel = this.InputChannel;
            _with1.InputDeviceName = this.InputDeviceName;
            _with1.LocalControl = this.LocalControl;
            _with1.NoteDuration = this.NoteDuration;
            _with1.OutputChannel = this.OutputChannel;
            _with1.OutputDeviceName = this.OutputDeviceName;
            _with1.PatchNumber = this.PatchNumber;
            _with1.SendPatchChangeOnOpen = this.SendPatchChangeOnOpen;
            _with1.Transpose = this.Transpose;
            _with1.Volume = this.Volume;
            return Data;
        }

        //-- Applies the settings from an InstrumentData object and
        //   opens the MIDI port
        public void SetSettings(InstrumentData Instrument)
        {
            bool Open = _Engaged;
            Close();
            var _with2 = Instrument;
            this.FilterAfterTouch = _with2.FilterAftertouch;
            this.InputChannel = _with2.InputChannel;
            this.InputDeviceName = _with2.InputDeviceName;
            this.LocalControl = _with2.LocalControl;
            this.NoteDuration = _with2.NoteDuration;
            this.OutputChannel = _with2.OutputChannel;
            this.OutputDeviceName = _with2.OutputDeviceName;
            this.PatchNumber = _with2.PatchNumber;
            this.Transpose = _with2.Transpose;
            this.SendPatchChangeOnOpen = _with2.SendPatchChangeOnOpen;
            this.Volume = _with2.Volume;
            if (Open == true)
            {
                try
                {
                    this.Open();
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        //-- Open both input and output devices
        public void Open()
        {
            if (_OutputDeviceID == -1 & _InputDeviceID == -1)
            {
                throw new Exception("You must specify an Input Device, Output Device, or both");
            }

            //-- Close the instrument if its already open
            try
            {
                if (_Engaged == true)
                {
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            //-- Open the new port
            try
            {
                //-- Only do this if we're not already engaged
                if (_Engaged == false)
                {
                    //-- Input Device
                    if (_InputDeviceID != -1)
                    {
                        hMidiIN = OpenMIDIInPort(ref _InputDeviceID, true);
                    }
                    //-- Output Device
                    if (_OutputDeviceID != -1)
                    {
                        hMidiOUT = OpenMIDIOutPort(ref _OutputDeviceID, true);
                    }
                    //-- Engage
                    _Engaged = true;
                    //-- Apply properties
                    LocalControl = _LocalControl;
                    if (_sendPatchChangeOnOpen == true)
                    {
                        PatchNumber = _PatchNumber;
                    }
                    Volume = _Volume;
                    //-- Write to trace log
                    Trace.WriteLine("------------------------------------");
                    Trace.WriteLine("Device Open at " + DateAndTime.Now.ToString());
                    Trace.Indent();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error opening MIDI port" + Constants.vbCrLf + ex.Message, ex);
            }
        }

        //-- Close both input and output devices
        public void Close()
        {
            try
            {
                MyNotesOff();
                if (_Engaged == true)
                {
                    //-- Send all notes off
                    AllNotesOff();
                    //-- Let the NoteDuration thread finish
                    if (_NoteDuration > 0)
                    {
                        ev.WaitOne();
                    }
                    //-- Close the input device
                    if (hMidiIN != 0)
                    {
                        hMidiIN = OpenMIDIInPort(ref _InputDeviceID, false);
                    }
                    //-- Close the output device
                    if (hMidiOUT != 0)
                    {
                        hMidiOUT = OpenMIDIOutPort(ref _OutputDeviceID, false);
                    }
                    //-- Set Engaged to False
                    _Engaged = false;

                    Trace.Unindent();
                    Trace.WriteLine("Device Closed at " + DateAndTime.Now.ToString());
                    Trace.WriteLine("------------------------------------");
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error closing MIDI Port" + Constants.vbCrLf + ex.Message, ex);
            }
        }

        //-- Sends MIDI Data
        public void Send(byte Channel, byte Status, byte Data1 = 0, byte Data2 = 0, byte Data3 = 0)
        {
            int MidiMsg = 0;
            byte CmdAndChannel = 0;
            CmdAndChannel = StuffByte(Status, Channel);
            MidiMsg = StuffInt32(CmdAndChannel, Data1, Data2, Data3);
            MIDI.midiOutShortMsg(hMidiOUT, MidiMsg);
        }

        //-- Send MIDI Data by specifying a message from an Enumeration
        public void SendMessage(MIDIStatusMessages Msg, ref byte Data1, byte Data2 = 0, byte Data3 = 0)
        {
            int MidiMsg = 0;
            byte CmdAndChannel = 0;
            CmdAndChannel = StuffByte(GetByte(Msg), OutputChannel);
            MidiMsg = StuffInt32(CmdAndChannel, Data1, Data2, Data3);
            MIDI.midiOutShortMsg(hMidiOUT, MidiMsg);
        }

        //-- Send a controller change by using an Enumeration
        public void SendControllerChange(MIDIControllers Controller, ref byte Data1, byte Data2 = 0)
        {
            int MidiMsg = 0;
            byte CmdAndChannel = 0;
            CmdAndChannel = StuffByte(GetByte(MIDIStatusMessages.ControllerChange), OutputChannel);
            MidiMsg = StuffInt32(CmdAndChannel, GetByte(Controller), Data1, Data2);
            MIDI.midiOutShortMsg(hMidiOUT, MidiMsg);
        }

        //-- Return a list of the input device names as a string array
        public static string[] InDeviceNames()
        {
            int num = 0;
            int i = 0;
            MIDI.MIDIINCAPS Caps = default(MIDI.MIDIINCAPS);
            string[] names = null;

            num = MIDI.midiInGetNumDevs();
            if (num > 0)
            {
                names = new string[num];
                for (i = 0; i <= num - 1; i++)
                {
                    MIDI.midiInGetDevCaps(i, ref Caps, Strings.Len(Caps));
                    names[i] = Caps.szPname;
                }
            }
            return names;
        }

        //-- Return a list of output device names as a string array
        public static string[] OutDeviceNames()
        {
            int num = 0;
            int i = 0;
            MIDI.MIDIOUTCAPS Caps = default(MIDI.MIDIOUTCAPS);
            string[] names = null;

            num = MIDI.midiOutGetNumDevs();
            if (num > 0)
            {
                names = new string[num];
                for (i = 0; i <= num - 1; i++)
                {
                    MIDI.midiOutGetDevCaps(i, ref Caps, Strings.Len(Caps));
                    names[i] = Caps.szPname;
                }
            }
            return names;
        }

        //-- Change the patch (output) by providing a GM Instrument name using an Enum
        public void ChangePatchGM(GMInstruments Instrument)
        {
            try
            {
                this.PatchNumber = GetByte(Instrument);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //-- Change the patch (output) by GM name
        public void ChangePatchGM(string GMPatchName)
        {
            try
            {
                this.PatchNumber = GetByte(Array.IndexOf(GMInstrumentNames, GMPatchName));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //-- Send the AllSoundOff (panic) message
        public void AllSoundOff()
        {
            byte Msg = 0;
            Int32 MidiMsg = default(Int32);
            if (_Engaged == false)
                return;
            try
            {
                Msg = StuffByte(GetByte(MIDIStatusMessages.ChannelModeMessage), _OutputChannel);
                MidiMsg = StuffInt32(Msg, 0x78, 0, 0);
                MIDI.midiOutShortMsg(hMidiOUT, MidiMsg);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //-- Reset all the output controllers
        public void ResetAllControllers()
        {
            byte Msg = 0;
            Int32 MidiMsg = default(Int32);
            if (_Engaged == false)
                return;
            try
            {
                Msg = StuffByte(GetByte(MIDIStatusMessages.ChannelModeMessage), _OutputChannel);
                MidiMsg = StuffInt32(Msg, 0x79, 0, 0);
                MIDI.midiOutShortMsg(hMidiOUT, MidiMsg);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //-- Turn off all notes currently playing
        public void AllNotesOff()
        {
            byte Msg = 0;
            Int32 MidiMsg = default(Int32);
            if (_Engaged == false)
                return;
            try
            {
                Msg = StuffByte(GetByte(MIDIStatusMessages.ChannelModeMessage), _OutputChannel);
                MidiMsg = StuffInt32(Msg, 0x7b, 0, 0);
                MIDI.midiOutShortMsg(hMidiOUT, MidiMsg);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //-- Play a note by number (0-127), specifying a velocity value (0-127)
        public void PlayNote(byte Note, byte Velocity)
        {
            byte Msg = 0;
            Int32 MidiMsg = default(Int32);

            if (_Engaged == false)
                return;
            Msg = StuffByte(GetByte(MIDIStatusMessages.NoteOn), _OutputChannel);
            MidiMsg = StuffInt32(Msg, GetByte(Note + _Transpose), Velocity, 0);
            try
            {
                MIDI.midiOutShortMsg(hMidiOUT, MidiMsg);
                if (_NoteDuration > 0)
                {
                    NotesOff no = default(NotesOff);
                    no.Note = GetByte(Note + _Transpose);
                    no.ShutoffTime = DateAndTime.DateAdd(DateInterval.Second, _NoteDuration, DateAndTime.Now);
                    lock (notesToTurnOff)
                    {
                        notesToTurnOff.Add(no);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //-- Stop playing a note by passing the note number (0-127) and optionally velocity (0-127)
        public void StopNote(byte Note, byte Velocity = 64)
        {
            byte Msg = 0;

            Int32 MidiMsg = default(Int32);

            if (_Engaged == false)
                return;
            Msg = StuffByte(GetByte(MIDIStatusMessages.NoteOff), _OutputChannel);
            MidiMsg = StuffInt32(Msg, GetByte(Note + _Transpose), Velocity, 0);
            try
            {
                MIDI.midiOutShortMsg(hMidiOUT, MidiMsg);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static byte GetByte(object obj)
        {
            byte b = 0;
            if (Information.IsNumeric(obj) == false)
            {
                b = 0;
            }
            else
            {
                switch (Convert.ToInt32(obj))
                {
                    // ERROR: Case labels with binary operators are unsupported : LessThan
                    case 0: 
                    
                        b = 0;
                        break;
                    // ERROR: Case labels with binary operators are unsupported : GreaterThan
                    case 255:
                        b = 255;
                        break;
                    default:
                        b = Convert.ToByte(obj);
                        break;
                }
            }
            return b;
        }

        #endregion

        #region " Property Handlers "

        //-- Boolean that the programmer can use for any reason.
        //   I use it to persist the state of a leslie simulator
        //   for Native Instruments' B4 Virtual Organ
        private bool _UserDefinedBool1;
        public bool UserDefinedBool1
        {
            get { return _UserDefinedBool1; }
            set { _UserDefinedBool1 = value; }
        }

        //-- Boolean that the programmer can use for any reason.
        private bool _UserDefinedBool2;
        public bool UserDefinedBool2
        {
            get { return _UserDefinedBool2; }
            set { _UserDefinedBool2 = value; }
        }

        //-- Setting this to a valid text file name turns on tracing.
        //   You can use this to monitor data and events as they 
        //   happen in real time.
        private string _OutputTraceLogFile;
        public string OutputTraceLogFile
        {
            get { return _OutputTraceLogFile; }
            set
            {
                try
                {
                    if (value == null)
                    {
                        if (Trace.Listeners.Count > 0)
                        {
                            Trace.Listeners[0].Close();
                            Trace.Listeners.Clear();
                        }
                        _OutputTraceLogFile = value;
                    }
                    else if (string.IsNullOrEmpty(value))
                    {
                        if (Trace.Listeners.Count > 0)
                        {
                            Trace.Listeners[0].Close();
                            Trace.Listeners.Clear();
                        }
                        _OutputTraceLogFile = value;
                    }
                    else
                    {
                        Trace.Listeners.Add(new TextWriterTraceListener(value));
                        Trace.AutoFlush = true;
                        _OutputTraceLogFile = value;
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        //-- When True, a patch change is sent when the output device is open
        private bool _sendPatchChangeOnOpen = false;
        public bool SendPatchChangeOnOpen
        {
            get { return _sendPatchChangeOnOpen; }
            set { _sendPatchChangeOnOpen = value; }
        }

        //-- Turns local control off and on
        private bool _LocalControl = true;
        public bool LocalControl
        {
            get { return _LocalControl; }
            set
            {
                if (_Engaged == true)
                {
                    byte Msg = 0;
                    Int32 MidiMsg = default(Int32);
                    try
                    {
                        Msg = StuffByte(GetByte(MIDIStatusMessages.ChannelModeMessage), _OutputChannel);
                        if (value == true)
                        {
                            MidiMsg = StuffInt32(Msg, 0x7a, 0x7f, 0);
                        }
                        else
                        {
                            MidiMsg = StuffInt32(Msg, 0x7a, 0, 0);
                        }
                        MIDI.midiOutShortMsg(hMidiOUT, MidiMsg);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Error setting Local Control: " + ex.Message, ex);
                    }
                }
                _LocalControl = value;
            }
        }

        //-- Set to a number of semitones less than or greater than zero
        //   to transpose output notes up or down
        private Int32 _Transpose = 0;
        public Int32 Transpose
        {
            get { return _Transpose; }
            set { _Transpose = value; }
        }

        //-- When True, aftertouch (channel pressure) on the input device
        //   is ignored
        private bool _FilterAfterTouch = true;
        public bool FilterAfterTouch
        {
            get { return _FilterAfterTouch; }
            set { _FilterAfterTouch = value; }
        }

        //-- When you set the NoteDuration property, you can just call 
        //   PlayNote and the notes will be cut off (NOTE_OFF sent) 
        //   after NoteDuration number of seconds
        private Int32 _NoteDuration = 0;
        public Int32 NoteDuration
        {
            get { return _NoteDuration; }
            set
            {
                if (value > 0)
                {
                    System.Threading.Thread T = new System.Threading.Thread(ShutoffNoteCallback);
                    T.Start();
                }
                _NoteDuration = value;
            }
        }
        //-- Event used to communicate from the ShutoffNoteCallback sub

        private System.Threading.AutoResetEvent ev = new System.Threading.AutoResetEvent(false);
        //-- Returns True if there are devices open
        public bool Engaged
        {
            get { return _Engaged; }
        }

        //-- Lets you specify an output device by name
        private string _OutputDeviceName = "";
        public string OutputDeviceName
        {
            get { return _OutputDeviceName; }
            set
            {
                if (value == null)
                    return;
                Int32 i = default(Int32);
                _OutputDeviceID = -1;
                string[] devices = OutDeviceNames();
                if ((devices != null))
                {
                    for (i = 0; i <= devices.Length - 1; i++)
                    {
                        if ((devices[i] != null))
                        {
                            if (devices[i].ToLower() == value.ToLower())
                            {
                                _OutputDeviceName = value;
                                _OutputDeviceID = i;
                                break; // TODO: might not be correct. Was : Exit For
                            }
                        }
                    }
                }
            }
        }

        //-- Lets you specify an input device by name
        private string _InputDeviceName = "";
        public string InputDeviceName
        {
            get { return _InputDeviceName; }
            set
            {
                if (value == null)
                    return;
                Int32 i = default(Int32);
                _InputDeviceID = -1;
                string[] devices = InDeviceNames();
                if ((devices != null))
                {
                    for (i = 0; i <= devices.Length - 1; i++)
                    {
                        if ((devices[i] != null))
                        {
                            if (devices[i].ToLower() == value.ToLower())
                            {
                                _InputDeviceName = value;
                                _InputDeviceID = i;
                                break; // TODO: might not be correct. Was : Exit For
                            }
                        }
                    }
                }
            }
        }

        //-- Specify the zero-based MIDI channel on the output device
        private byte _OutputChannel = 0;
        public byte OutputChannel
        {
            get { return _OutputChannel; }
            set { _OutputChannel = value; }
        }

        //-- Specify the zero-based MIDI channel on the input device
        private byte _InputChannel = 0;
        public byte InputChannel
        {
            get { return _InputChannel; }
            set { _InputChannel = value; }
        }

        //-- Sets the volume of the output device
        private byte _Volume = 127;
        public byte Volume
        {
            get { return _Volume; }

            set
            {
                _Volume = value;
                if (_Engaged == true)
                {
                    try
                    {
                        byte Msg = 0;
                        Int32 MidiMsg = default(Int32);
                        Msg = StuffByte(GetByte(MIDIStatusMessages.ControllerChange), _OutputChannel);
                        MidiMsg = StuffInt32(Msg, 7, value, 0);
                        MIDI.midiOutShortMsg(hMidiOUT, MidiMsg);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Error setting Volume: " + ex.Message, ex);
                    }
                }
            }
        }

        //-- Sets the patch number for the output device
        private byte _PatchNumber = 0;
        public byte PatchNumber
        {
            get { return _PatchNumber; }
            set
            {
                _PatchNumber = value;
                if (_Engaged == true)
                {
                    try
                    {
                        byte Msg = 0;
                        Int32 MidiMsg = default(Int32);
                        Msg = StuffByte(GetByte(MIDIStatusMessages.ProgramChange), _OutputChannel);
                        MidiMsg = StuffInt32(Msg, value, 0, 0);
                        MIDI.midiOutShortMsg(hMidiOUT, MidiMsg);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
        }

        #endregion

   private byte MZAHL = 2;
   private byte RAUSCHEN= 2;

    private long lSumme;  // Summe von Mittelwerten
    private int iAnzahl;
    private long lSummeGes;
    private long AltWert;

// Messung des elektrischen Feldes
byte getIntegralValue(int MessWert)
{
	long lVal = MessWert;
	long lRet=0;
	lVal=255-lVal;
	iAnzahl++;
	lSumme +=lVal;
	if (iAnzahl > MZAHL)
	{
		lRet = lSumme/iAnzahl;
		iAnzahl = 0;
		lSumme = 0;
		long lDiff = Math.Abs(lRet-AltWert);
		AltWert=lRet;
		if ((lRet > RAUSCHEN)&&(lDiff>0))
		{
			if (lRet > 127)
				lRet=127;
			if (lRet < 4)
				lRet=0;
		}
		else
		{
			lRet=0L;
		}
	}
	return (byte)lRet;
}


private long lSumme2=0;  // Summe von Mittelwerten
private int iAnzahl2=0;
private long AltWert2=0;

// IR-Entfernungsmessung
byte getSmoothedValue(int MessWert)
{
	long lRet=0L;
	long lVal = MessWert;
	iAnzahl2++;
	lSumme2 +=lVal;
	if (iAnzahl2 > MZAHL)
	{
		lRet = lSumme2/iAnzahl2;
		iAnzahl2 = 0;
		lSumme2 = 0;
		long lDiff = Math.Abs(lRet-AltWert2);
		AltWert2=lRet;
		if ((lRet > RAUSCHEN)&&(lDiff>0))
		{
			lRet-=38L;
			if (lRet < 0L)
				lRet=0L;
			if (lRet >127L)
				lRet =127L;
		}
		else
		{
			lRet=0L;
		}
	}
	return (byte)lRet;
}

long getMeanValue(byte MessWert)
{
	long lVal = (long)MessWert;
	iAnzahl++;
	lSumme +=lVal;
	if (iAnzahl<MZAHL)
	{
		return -1;
	}
	else
	{
		long lRet =lSumme/iAnzahl;     // eine Messreihe ist rum
		iAnzahl = 0;
		lSumme = 0;
		if (lRet > 0)
		{
			lSummeGes=(lSummeGes+lRet)/2;    // Messreihe zur vorherigen hinzufügen
			return -1;
		}
		else

		{
			long lRet2=lSummeGes;		// Gesamtergebnis ausgeben
			lSummeGes=0;
			return lRet2;
		}
	}
}



public byte GetMidiNote(byte Value)    // momentan Wert zwischen  20 und 200
{
    if (Value > 0)
    {
        int val = Value/5;  // 4 und 40
        int oktave = val;
        oktave = oktave / NOTE_ANZ;   // Oktave
        if (oktave > 7)
            oktave = 7;
        if (oktave < 0)
            oktave = 0;
        int index = val % NOTE_ANZ;
        byte ucNote = (byte)(whTonleiter[index] + ((oktave + 1) * 12));
        return ucNote;
    }
    return 0;
}

public void MyNotesOff()
{
    for (int i = 0; i < 128; i++)
    {
        Send(0, (byte)MIDIStatusMessages.NoteOn, (byte)i, 0);
        Send(1, (byte)MIDIStatusMessages.NoteOn, (byte)i, 0);
        Send(0, (byte)MIDIStatusMessages.NoteOff, (byte)i, 0);
        Send(1, (byte)MIDIStatusMessages.NoteOff, (byte)i, 0);
    }
}

//DateTime OldDtLeft;
//DateTime OldDtRight;

public void TranslateToNote(byte chan, byte NoteLeft, byte LastNoteLeft, byte NoteRight, byte LastNoteRight)
{
//    DateTime dt = DateTime.Now;
    if (LastNoteLeft != NoteLeft)
    {
      //  if ((dt - OldDtLeft) > TimeSpan.FromSeconds(1.0))
        {
            if (LastNoteLeft > 0)
                Send(chan, (byte)MIDIStatusMessages.NoteOff, LastNoteLeft, 0);
            if (NoteLeft > 0)
                Send(chan, (byte)MIDIStatusMessages.NoteOn, NoteLeft, 100);
     //       OldDtLeft = dt;
            System.Console.WriteLine("Linke Note {0:N}", NoteLeft);

        }
    }
    if (LastNoteRight != NoteRight)
    {
     //   if ((dt - OldDtRight) > TimeSpan.FromSeconds(1.0))
        {
            chan++;
            if (LastNoteRight > 0)
                Send(chan, (byte)MIDIStatusMessages.NoteOff, LastNoteRight, 0);
            if (NoteRight > 0)
                Send(chan, (byte)MIDIStatusMessages.NoteOn, NoteRight, 100);
       //     OldDtRight = dt;
            System.Console.WriteLine("Rechte Note {0:N}", NoteRight);
        }
    }
}

        #region " Shared Arrays "

        //-- General MIDI Instrument Names
        public static string[] GMInstrumentNames = {
		"AcousticGrandPiano",
		"BrightAcousticPiano",
		"ElectricGrandPiano",
		"HonkyTonkPiano",
		"ElectricPiano1",
		"ElectricPiano2",
		"Harpsichord",
		"Clavi",
		"Celesta",
		"Glockenspiel",
		"MusicBox",
		"Vibraphone",
		"Marimba",
		"Xylophone",
		"TubularBells",
		"Dulcimer",
		"DrawbarOrgan",
		"PercussiveOrgan",
		"RockOrgan",
		"ChurchOrgan",
		"ReedOrgan",
		"Accordion",
		"Harmonica",
		"TangoAccordion",
		"AcousticGuitarNylon",
		"AcousticGuitarSteel",
		"ElectricGuitarJazz",
		"ElectricGuitarClean",
		"ElectricGuitarMuted",
		"OverdrivenGuitar",
		"DistortionGuitar",
		"GuitarHarmonics",
		"AcousticBass",
		"ElectricBassFinger",
		"ElectricBassPick",
		"FretlessBass",
		"SlapBass1",
		"SlapBass2",
		"SynthBass1",
		"SynthBass2",
		"Violin",
		"Viola",
		"Cello",
		"Contrabass",
		"TremoloStrings",
		"PizzicatoStrings",
		"OrchestralHarp",
		"Timpani",
		"StringEnsemble1",
		"StringEnsemble2",
		"SynthStrings1",
		"SynthStrings2",
		"ChoirAahs",
		"VoiceOohs",
		"SynthVoice",
		"OrchestraHit",
		"Trumpet",
		"Trombone",
		"Tuba",
		"MutedTrumpet",
		"FrenchHorn",
		"BrassSection",
		"SynthBrass1",
		"SynthBrass2",
		"SopranoSax",
		"AltoSax",
		"TenorSax",
		"BaritoneSax",
		"Oboe",
		"EnglishHorn",
		"Bassoon",
		"Clarinet",
		"Piccolo",
		"Flute",
		"Recorder",
		"PanFlute",
		"BlownBottle",
		"Shakuhachi",
		"Whistle",
		"Ocarina",
		"Lead1Square",
		"Lead2Sawtooth",
		"Lead3Calliope",
		"Lead4Chiff",
		"Lead5Charang",
		"Lead6Voice",
		"Lead7Fifths",
		"Lead8BassPluslead",
		"Pad1NewAge",
		"Pad2Warm",
		"Pad3Polysynth",
		"Pad4Choir",
		"Pad5Bowed",
		"Pad6Metallic",
		"Pad7Halo",
		"Pad8Sweep",
		"FX1Rain",
		"FX2Soundtrack",
		"FX3Crystal",
		"FX4Atmosphere",
		"FX5Brightness",
		"FX6Goblins",
		"FX7Echoes",
		"FX8SciFi",
		"Sitar",
		"Banjo",
		"Shamisen",
		"Koto",
		"Kalimba",
		"Bagpipe",
		"Fiddle",
		"Shanai",
		"TinkleBell",
		"Agogo",
		"SteelDrums",
		"Woodblock",
		"TaikoDrum",
		"MelodicTom",
		"SynthDrum",
		"ReverseCymbal",
		"GuitarFretNoise",
		"BreathNoise",
		"Seashore",
		"BirdTweet",
		"TelephoneRing",
		"Helicopter",
		"Applause",
		"Gunshot"

	};
        #endregion

    }

    #region " Enums "

    public enum MIDIStatusMessages
    {
        NoteOff = (0x8),
        NoteOn = (0x9),
        PolyphonicKeyPressure = (0xa),
        ControllerChange = (0xb),
        ProgramChange = (0xc),
        ChannelPressure = (0xd),
        PitchBend = (0xe),
        ChannelModeMessage = (0xb)
    }

    public enum MIDIControllers
    {
        ModWheel = 1,
        BreathController = 2,
        FootController = 4,
        PortamentoTime = 5,
        MainVolume = 7,
        Balance = 8,
        Pan = 10,
        ExpressionController = 11,
        DamperPedal = 64,
        Portamento = 65,
        Sostenuto = 66,
        SoftPedal = 67,
        Hold2 = 69,
        ExternalEffectsDepth = 91,
        TremeloDepth = 92,
        ChorusDepth = 93,
        DetuneDepth = 94,
        PhaserDepth = 95,
        DataIncrement = 96,
        DataDecrement = 97,
        AllNotesOff = 123
    }

    public enum GMInstruments
    {
        AcousticGrandPiano,
        BrightAcousticPiano,
        ElectricGrandPiano,
        HonkyTonkPiano,
        ElectricPiano1,
        ElectricPiano2,
        Harpsichord,
        Clavi,
        Celesta,
        Glockenspiel,
        MusicBox,
        Vibraphone,
        Marimba,
        Xylophone,
        TubularBells,
        Dulcimer,
        DrawbarOrgan,
        PercussiveOrgan,
        RockOrgan,
        ChurchOrgan,
        ReedOrgan,
        Accordion,
        Harmonica,
        TangoAccordion,
        AcousticGuitarNylon,
        AcousticGuitarSteel,
        ElectricGuitarJazz,
        ElectricGuitarClean,
        ElectricGuitarMuted,
        OverdrivenGuitar,
        DistortionGuitar,
        GuitarHarmonics,
        AcousticBass,
        ElectricBassFinger,
        ElectricBassPick,
        FretlessBass,
        SlapBass1,
        SlapBass2,
        SynthBass1,
        SynthBass2,
        Violin,
        Viola,
        Cello,
        Contrabass,
        TremoloStrings,
        PizzicatoStrings,
        OrchestralHarp,
        Timpani,
        StringEnsemble1,
        StringEnsemble2,
        SynthStrings1,
        SynthStrings2,
        ChoirAahs,
        VoiceOohs,
        SynthVoice,
        OrchestraHit,
        Trumpet,
        Trombone,
        Tuba,
        MutedTrumpet,
        FrenchHorn,
        BrassSection,
        SynthBrass1,
        SynthBrass2,
        SopranoSax,
        AltoSax,
        TenorSax,
        BaritoneSax,
        Oboe,
        EnglishHorn,
        Bassoon,
        Clarinet,
        Piccolo,
        Flute,
        Recorder,
        PanFlute,
        BlownBottle,
        Shakuhachi,
        Whistle,
        Ocarina,
        Lead1Square,
        Lead2Sawtooth,
        Lead3Calliope,
        Lead4Chiff,
        Lead5Charang,
        Lead6Voice,
        Lead7Fifths,
        Lead8BassPluslead,
        Pad1NewAge,
        Pad2Warm,
        Pad3Polysynth,
        Pad4Choir,
        Pad5Bowed,
        Pad6Metallic,
        Pad7Halo,
        Pad8Sweep,
        FX1Rain,
        FX2Soundtrack,
        FX3Crystal,
        FX4Atmosphere,
        FX5Brightness,
        FX6Goblins,
        FX7Echoes,
        FX8SciFi,
        Sitar,
        Banjo,
        Shamisen,
        Koto,
        Kalimba,
        Bagpipe,
        Fiddle,
        Shanai,
        TinkleBell,
        Agogo,
        SteelDrums,
        Woodblock,
        TaikoDrum,
        MelodicTom,
        SynthDrum,
        ReverseCymbal,
        GuitarFretNoise,
        BreathNoise,
        Seashore,
        BirdTweet,
        TelephoneRing,
        Helicopter,
        Applause,
        Gunshot
    }

    #endregion

    [Serializable()]
    public class InstrumentData
    {
        public string Name;
        public string InputDeviceName;
        public string OutputDeviceName;
        public byte InputChannel;
        public byte OutputChannel;
        public Int32 NoteDuration;
        public bool SendPatchChangeOnOpen;
        public bool LocalControl;
        public bool FilterAftertouch = true;
        public Int32 Transpose;
        public byte PatchNumber;
        public byte Volume = 127;
    }
}
