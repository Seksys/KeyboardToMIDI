using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Threading;

namespace libMIDI
{
    /// <summary>
    /// MIDI class for playing MIDI files and notes.
    /// </summary>
    public class MIDI
    {
        public delegate void MIDIInstrumentChangedEvent(MIDI_Instruments inst, string name);
        public event MIDIInstrumentChangedEvent OnMidiInstrumentChanged;

        #region PInvoke
        [DllImport("winmm.dll")]
        private static extern long mciSendString(string command, StringBuilder returnValue, int returnLength, IntPtr winHandle);

        [DllImport("winmm.dll")]
        private static extern int midiOutGetNumDevs();

        [DllImport("winmm.dll")]
        private static extern int midiOutGetDevCaps(Int32 uDeviceID, ref MidiOutCaps lpMidiOutCaps, UInt32 cbMidiOutCaps);

        [DllImport("winmm.dll")]
        private static extern int midiOutOpen(ref IntPtr handle, int deviceID, MidiCallBack proc, int instance, int flags);

        [DllImport("winmm.dll")]
        protected static extern int midiOutShortMsg(IntPtr handle, int message);

        [DllImport("winmm.dll")]
        protected static extern int midiOutClose(IntPtr handle);
        #endregion

        private delegate void MidiCallBack(int handle, int msg, int instance, int param1, int param2);

        private IntPtr hndMIDI;
        private int cntMIDIDev;

        public bool IsOpen { internal set; get; }
        public int SelectedDevice { internal set; get; }
        public MIDI_Instruments SelectedInstrument { internal set; get; }

        public MIDI()
        {
            cntMIDIDev = midiOutGetNumDevs();
            SelectedDevice = 0;

            hndMIDI = IntPtr.Zero;
            Console.WriteLine("MIDI -> Devices:{0}", cntMIDIDev);
        }
        public bool Open() { return Open(0); }
        public bool Open(int devIndex)
        {

            //var numDevs = midiOutGetNumDevs();
            if (devIndex <= cntMIDIDev)
            {
                SelectedDevice = devIndex;
                MidiOutCaps myCaps = new MidiOutCaps();
                int moCaps = midiOutGetDevCaps(devIndex, ref myCaps, (UInt32)Marshal.SizeOf(myCaps));
                int moOpen = midiOutOpen(ref hndMIDI, devIndex, null, 0, 0);

                Console.WriteLine("MIDI -> Open() -> Device:{0}, Caps:{1}, Open:{2}, Name:{3} - v{4}", devIndex, moCaps == 0?"OK":"ER", moOpen == 0?"OK":"ER", myCaps.szPname, myCaps.vDriverVersion);
                IsOpen = true;

                SetInstrument(MIDI_Instruments.AcousticGrandPiano);

                return IsOpen;
            }
            else {
                IsOpen = false;
                return IsOpen;
            }
        }
        public void Close() { if (IsOpen) { int res = midiOutClose(hndMIDI); IsOpen = false; } }

        /// <summary>
        /// Sends a command via the MCI interface.
        /// </summary>
        /// <param name="command">Command to send.</param>
        /// <returns>Status result.</returns>
        static string MCI(string command)
        {
            int returnLength = 256;
            StringBuilder reply = new StringBuilder(returnLength);
            mciSendString(command, reply, returnLength, IntPtr.Zero);
            return reply.ToString();
        }


        private static byte EncodeNote(string note, int octave)
        {
            //12 notes per octave
            //-1 to 9 octaves, 11 total
            //128 possible notes, drop off is at "G-9", any notes past this one cant be played, too high.
            byte n = 0; //0 defines a rest for our purposes
            switch (note)
            {
                case "B#": n = 1; break;
                case "C-": n = 1; break;
                case "C#": n = 2; break;
                case "Db": n = 2; break;
                case "D-": n = 3; break;
                case "D#": n = 4; break;
                case "Eb": n = 4; break;
                case "E-": n = 5; break;
                case "Fb": n = 5; break;
                case "E#": n = 6; break;
                case "F-": n = 6; break;
                case "F#": n = 7; break;
                case "Gb": n = 7; break;
                case "G-": n = 8; break;
                case "G#": n = 9; break;
                case "Ab": n = 9; break;
                case "A-": n = 10; break;
                case "A#": n = 11; break;
                case "Bb": n = 11; break;
                case "B-": n = 12; break;
                case "Cb": n = 12; break;
                default: break;
            }

            if (n != 0)
            {
                n--;    //decrement to adjust "rest" position
                byte o = 0;
                switch (octave)
                {
                    case 9: o = 10; break;
                    case 8: o = 9; break;
                    case 7: o = 8; break;
                    case 6: o = 7; break;
                    case 5: o = 6; break;
                    case 4: o = 5; break;
                    case 3: o = 4; break;
                    case 2: o = 3; break;
                    case 1: o = 2; break;
                    case 0: o = 1; break;
                    case -1: o = 0; break;
                    default: break;
                }
                //Octave adjustments for when you cross an octave boundary (B to C Threshold)
                if (note == "Cb") { o--; }
                if (note == "B#") { o++; }
                return (byte)((n + (o * 12)));


            }
            else {
                return 0;
            }

        }
        private static int SetNoteCmdByte(byte cmd, string note, int octave, byte vel = 0x7F) { return (vel << 16) + (EncodeNote(note, octave) << 8) + cmd; }


        public void PlayFile(string filePath,int durationMS = 5000, string alias = "midiFileMusic")
        {
            string res = "";
            res = MCI("open \"" + filePath + "\" alias " + alias);
            res = MCI("play midiMusic");
            Pause(durationMS);
            res = MCI("close midiMusic");
        }
        public void PlayNote(string note, int octave, byte velocity = 0x7F) { if (IsOpen) { if (note != "R-") { int res = midiOutShortMsg(hndMIDI, SetNoteCmdByte((byte)MIDICommands.NoteDn, note, octave, velocity)); if (res != 0) { Console.WriteLine("MIDI -> NoteDown Error: {0}", (MMRESULT)res); } } } }
        public void StopNote(string note, int octave, byte velocity = 0x7F) { if (IsOpen) { if (note != "R-") { int res = midiOutShortMsg(hndMIDI, SetNoteCmdByte((byte)MIDICommands.NoteUp, note, octave, velocity)); if (res != 0) { Console.WriteLine("MIDI -> NoteUp   Error: {0}", (MMRESULT)res); } } } }

        public void PlayNote(ISimpleNote note) { PlayNote(note.Pitch, note.Octave, note.Velocity); }
        public void StopNote(ISimpleNote note) { StopNote(note.Pitch, note.Octave, note.Velocity); }

        public void SetInstrument(MIDI_Instruments instrument)
        {
            int res = 0;
            if (IsOpen)
            {
                if ((int)instrument == 31) { instrument = (MIDI_Instruments)30; } //Force not to use distortion guitar, sounds funny.
                SelectedInstrument = instrument;
                byte cmd = 0xC0;

                int msg = ((byte)instrument << 8) + cmd;
                res = midiOutShortMsg(hndMIDI, msg);
                if (res != 0 && OnMidiInstrumentChanged != null) { OnMidiInstrumentChanged(instrument, instrument.ToString()); }
                Console.WriteLine("MIDI -> SetInstrument() -> MIDI_Instrument.{0}", instrument);
            }
        }
        public static string GetInstrumentName(byte instrument)
        {
            string output = "";

            switch (instrument)
            {

                case 1: return "Acoustic Grand Piano";
                case 2: return "Bright Acoustic Piano";
                case 3: return "Electric Grand Piano";
                case 4: return "Honky - tonk Piano";
                case 5: return "Electric Piano 1";
                case 6: return "Electric Piano 2";
                case 7: return "Harpsichord";
                case 8: return "Clavi";
                case 9: return "Celesta";
                case 10: return "Glockenspiel";
                case 11: return "Music Box";
                case 12: return "Vibraphone";
                case 13: return "Marimba";
                case 14: return "Xylophone";
                case 15: return "Tubular Bells";
                case 16: return "Dulcimer";
                case 17: return "Drawbar Organ";
                case 18: return "Percussive Organ";
                case 19: return "Rock Organ";
                case 20: return "Church Organ";
                case 21: return "Reed Organ";
                case 22: return "Accordion";
                case 23: return "Harmonica";
                case 24: return "Tango Accordion";
                case 25: return "Acoustic Guitar(nylon)";
                case 26: return "Acoustic Guitar(steel)";
                case 27: return "Electric Guitar(jazz)";
                case 28: return "Electric Guitar(clean)";
                case 29: return "Electric Guitar(muted)";
                case 30: return "Overdriven Guitar";
                case 31: return "Distortion Guitar";
                case 32: return "Guitar harmonics";
                case 33: return "Acoustic Bass";
                case 34: return "Electric Bass(finger)";
                case 35: return "Electric Bass(pick)";
                case 36: return "Fretless Bass";
                case 37: return "Slap Bass 1";
                case 38: return "Slap Bass 2";
                case 39: return "Synth Bass 1";
                case 40: return "Synth Bass 2";
                case 41: return "Violin";
                case 42: return "Viola";
                case 43: return "Cello";
                case 44: return "Contrabass";
                case 45: return "Tremolo Strings";
                case 46: return "Pizzicato Strings";
                case 47: return "Orchestral Harp";
                case 48: return "Timpani";
                case 49: return "String Ensemble 1";
                case 50: return "String Ensemble 2";
                case 51: return "SynthStrings 1";
                case 52: return "SynthStrings 2";
                case 53: return "Choir Aahs";
                case 54: return "Voice Oohs";
                case 55: return "Synth Voice";
                case 56: return "Orchestra Hit";
                case 57: return "Trumpet";
                case 58: return "Trombone";
                case 59: return "Tuba";
                case 60: return "Muted Trumpet";
                case 61: return "French Horn";
                case 62: return "Brass Section";
                case 63: return "SynthBrass 1";
                case 64: return "SynthBrass 2";
                case 65: return "Soprano Sax";
                case 66: return "Alto Sax";
                case 67: return "Tenor Sax";
                case 68: return "Baritone Sax";
                case 69: return "Oboe";
                case 70: return "English Horn";
                case 71: return "Bassoon";
                case 72: return "Clarinet";
                case 73: return "Piccolo";
                case 74: return "Flute";
                case 75: return "Recorder";
                case 76: return "Pan Flute";
                case 77: return "Blown Bottle";
                case 78: return "Shakuhachi";
                case 79: return "Whistle";
                case 80: return "Ocarina";
                case 81: return "Lead 1(square)";
                case 82: return "Lead 2(sawtooth)";
                case 83: return "Lead 3(calliope)";
                case 84: return "Lead 4(chiff)";
                case 85: return "Lead 5(charang)";
                case 86: return "Lead 6(voice)";
                case 87: return "Lead 7(fifths)";
                case 88: return "Lead 8(bass + lead)";
                case 89: return "Pad 1(new age)";
                case 90: return "Pad 2(warm)";
                case 91: return "Pad 3(polysynth)";
                case 92: return "Pad 4(choir)";
                case 93: return "Pad 5(bowed)";
                case 94: return "Pad 6(metallic)";
                case 95: return "Pad 7(halo)";
                case 96: return "Pad 8(sweep)";
                case 97: return "FX 1(rain)";
                case 98: return "FX 2(soundtrack)";
                case 99: return "FX 3(crystal)";
                case 100: return "FX 4(atmosphere)";
                case 101: return "FX 5(brightness)";
                case 102: return "FX 6(goblins)";
                case 103: return "FX 7(echoes)";
                case 104: return "FX 8(sci - fi)";
                case 105: return "Sitar";
                case 106: return "Banjo";
                case 107: return "Shamisen";
                case 108: return "Koto";
                case 109: return "Kalimba";
                case 110: return "Bag pipe";
                case 111: return "Fiddle";
                case 112: return "Shanai";
                case 113: return "Tinkle Bell";
                case 114: return "Agogo";
                case 115: return "Steel Drums";
                case 116: return "Woodblock";
                case 117: return "Taiko Drum";
                case 118: return "Melodic Tom";
                case 119: return "Synth Drum";
                case 120: return "Reverse Cymbal";
                case 121: return "Guitar Fret Noise";
                case 122: return "Breath Noise";
                case 123: return "Seashore";
                case 124: return "Bird Tweet";
                case 125: return "Telephone Ring";
                case 126: return "Helicopter";
                case 127: return "Applause";
                case 128: return "Gunshot";
                default: output = "Unknown"; break;
            }


            return output;
        }


        /// <summary>
        /// Thread-safe Application.DoEvents()
        /// </summary>
        public static void DoEvents()
        {
            DispatcherFrame frame = new DispatcherFrame();
            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, new DispatcherOperationCallback(ExitFrame), frame);

            //see if a frame can just be pushed over and over.
            Dispatcher.PushFrame(frame);
        }
        private static object ExitFrame(object f)
        {
            ((DispatcherFrame)f).Continue = false;

            return null;
        }

        /// <summary>
        /// Thread-safe non-locking System.Threading.Thread.Sleep(duration). Program will resume when background task has been used for durationMS.
        /// </summary>
        /// <param name="durationMS">Duration of pause in MS.</param>
        /// <param name="useSafeMethod">If false will use program locking Sleep().</param>
        /// <returns></returns>
        public static DateTime Pause(int durationMS, bool useSafeMethod = true)
        {
            System.DateTime ThisMoment = System.DateTime.Now;
            System.TimeSpan duration = new System.TimeSpan(0, 0, 0, 0, durationMS);
            System.DateTime AfterWards = ThisMoment.Add(duration);

            if (!useSafeMethod) {
                System.Threading.Thread.Sleep(duration);
            } else {
                while (AfterWards >= ThisMoment) {
                    //See if moving this before the loop has the same effect.
                    DoEvents();
                    ThisMoment = System.DateTime.Now;
                }
            }
            return System.DateTime.Now;
        }

        public enum MIDI_Instruments
        {
            Null = 0,
            AcousticGrandPiano = 1,
            BrightAcousticPiano = 2,
            ElectricGrandPiano = 3,
            HonkytonkPiano = 4,
            ElectricPiano1 = 5,
            ElectricPiano2 = 6,
            Harpsichord = 7,
            Clavi = 8,
            Celesta = 9,
            Glockenspiel = 10,
            MusicBox = 11,
            Vibraphone = 12,
            Marimba = 13,
            Xylophone = 14,
            TubularBells =15,
            Dulcimer = 16,
            DrawbarOrgan = 17,
            PercussiveOrgan = 18,
            RockOrgan = 19,
            ChurchOrgan = 20,
            ReedOrgan = 21,
            Accordion = 22,
            Harmonica = 23,
            TangoAccordion = 24,
            AcousticGuitar_nylon = 25,
            AcousticGuitar_steel = 26,
            ElectricGuitar_jazz = 27,
            ElectricGuitar_clean = 28,
            ElectricGuitar_muted = 29,
            OverdrivenGuitar = 30,
            DistortionGuitar = 31,
            Guitarharmonics = 32,
            AcousticBass = 33,
            ElectricBass_finger = 34,
            ElectricBass_pick = 35,
            FretlessBass = 36,
            SlapBass1 = 37,
            SlapBass2 = 38,
            SynthBass1 = 39,
            SynthBass2 = 40,
            Violin = 41,
            Viola = 42,
            Cello = 43,
            Contrabass = 44,
            TremoloStrings = 45,
            PizzicatoStrings = 46,
            OrchestralHarp = 47,
            Timpani = 48,
            StringEnsemble1 = 49,
            StringEnsemble2 = 50,
            SynthStrings1 = 51,
            SynthStrings2 = 52,
            ChoirAahs = 53,
            VoiceOohs = 54,
            SynthVoice = 55,
            OrchestraHit = 56,
            Trumpet = 57,
            Trombone = 58,
            Tuba = 59,
            MutedTrumpet = 60,
            FrenchHorn = 61,
            BrassSection = 62,
            SynthBrass1 = 63,
            SynthBrass2 = 64,
            SopranoSax = 65,
            AltoSax = 66,
            TenorSax = 67,
            BaritoneSax = 68,
            Oboe = 69,
            EnglishHorn = 70,
            Bassoon = 71,
            Clarinet = 72,
            Piccolo = 73,
            Flute = 74,
            Recorder = 75,
            PanFlute = 76,
            BlownBottle = 77,
            Shakuhachi = 78,
            Whistle = 79,
            Ocarina = 80,
            Lead1_square = 81,
            Lead2_sawtooth = 82,
            Lead3_calliope = 83,
            Lead4_chiff = 84,
            Lead5_charang = 85,
            Lead6_voice = 86,
            Lead7_fifths = 87,
            Lead8_bass_lead = 88,
            Pad1_newage = 89,
            Pad2_warm = 90,
            Pad3_polysynth = 91,
            Pad4_choir = 92,
            Pad5_bowed = 93,
            Pad6_metallic = 94,
            Pad7_halo = 95,
            Pad8_sweep = 96,
            FX1_rain = 97,
            FX2_soundtrack = 98,
            FX3_crystal = 99,
            FX4_atmosphere = 100,
            FX5_brightness = 101,
            FX6_goblins = 102,
            FX7_echoes = 103,
            FX8_sci_fi =104,
            Sitar = 105,
            Banjo = 106,
            Shamisen = 107,
            Koto = 108,
            Kalimba = 109,
            Bagpipe = 110,
            Fiddle = 111,
            Shanai = 112,
            TinkleBell = 113,
            Agogo = 114,
            SteelDrums = 115,
            Woodblock = 117,
            TaikoDrum = 118,
            MelodicTom = 119,
            SynthDrum = 120,
            ReverseCymbal = 121,
            GuitarFretNoise = 122,
            BreathNoise = 123,
            Seashore = 124,
            BirdTweet = 125,
            TelephoneRing = 125,
            Helicopter = 126,
            Applause = 127,
            Gunshot = 128,
        }
        public enum MIDICommands {
            NoteDn = 0x90,
            NoteUp = 0x80               
        }
        public enum MMRESULT : uint
        {
            MMSYSERR_NOERROR = 0,
            MMSYSERR_ERROR = 1,
            MMSYSERR_BADDEVICEID = 2,
            MMSYSERR_NOTENABLED = 3,
            MMSYSERR_ALLOCATED = 4,
            MMSYSERR_INVALHANDLE = 5,
            MMSYSERR_NODRIVER = 6,
            MMSYSERR_NOMEM = 7,
            MMSYSERR_NOTSUPPORTED = 8,
            MMSYSERR_BADERRNUM = 9,
            MMSYSERR_INVALFLAG = 10,
            MMSYSERR_INVALPARAM = 11,
            MMSYSERR_HANDLEBUSY = 12,
            MMSYSERR_INVALIDALIAS = 13,
            MMSYSERR_BADDB = 14,
            MMSYSERR_KEYNOTFOUND = 15,
            MMSYSERR_READERROR = 16,
            MMSYSERR_WRITEERROR = 17,
            MMSYSERR_DELETEERROR = 18,
            MMSYSERR_VALNOTFOUND = 19,
            MMSYSERR_NODRIVERCB = 20,
            WAVERR_BADFORMAT = 32,
            WAVERR_STILLPLAYING = 33,
            WAVERR_UNPREPARED = 34
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MidiOutCaps
        {
            public UInt16 wMid;
            public UInt16 wPid;
            public UInt32 vDriverVersion;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public String szPname;

            public UInt16 wTechnology;
            public UInt16 wVoices;
            public UInt16 wNotes;
            public UInt16 wChannelMask;
            public UInt32 dwSupport;
        }


    }

}
