using libMIDI;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Input;

namespace wpfUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// MIDI Instance
        /// </summary>
        private MIDI midi;
        /// <summary>
        /// Which Keys are currently down, which allows holding notes.
        /// </summary>
        private List<Key> KeysDown = new List<Key>();

        #region Form Init and Dispose
        public MainWindow()
        {
            InitializeComponent();
            KeysDown = new List<Key>();
            midi = new MIDI();
            midi.Open();

            this.KeyDown += new KeyEventHandler(wpfKeyDown);
            this.KeyUp += new KeyEventHandler(wpfKeyUp);
        }

        /// <summary>
        /// When disposing be sure to close the MIDI, it checks itself to skip if its already closed.
        /// </summary>
        ~MainWindow() { midi.Close(); }
        #endregion


        /// <summary>
        /// This is where you map which keyboard keys play which note.
        /// </summary>
        /// <param name="k"></param>
        /// <returns></returns>
        public SimpleNote ConvertKeyToNote(Key k) {
            SimpleNote note = new SimpleNote();
            switch (k) {
                case Key.Z:  note =  new SimpleNote("C-", 3); break;
                case Key.S:  note =  new SimpleNote("C#", 3); break;
                case Key.X:  note =  new SimpleNote("D-", 3); break;
                case Key.D:  note =  new SimpleNote("D#", 3); break;
                case Key.C:  note =  new SimpleNote("E-", 3); break;
                case Key.V:  note =  new SimpleNote("F-", 3); break;
                case Key.G:  note =  new SimpleNote("F#", 3); break;
                case Key.B:  note =  new SimpleNote("G-", 3); break;
                case Key.H:  note =  new SimpleNote("G#", 3); break;
                case Key.N:  note =  new SimpleNote("A-", 3); break;
                case Key.J:  note =  new SimpleNote("A#", 3); break;
                case Key.M:  note =  new SimpleNote("B-", 3); break;
                //Next Octave
                case Key.Q:  note =  new SimpleNote("C-", 4); break;
                case Key.D2: note =  new SimpleNote("C#", 4); break;
                case Key.W:  note =  new SimpleNote("D-", 4); break;
                case Key.D3: note =  new SimpleNote("D#", 4); break;
                case Key.E:  note =  new SimpleNote("E-", 4); break;
                case Key.R:  note =  new SimpleNote("F-", 4); break;
                case Key.D5: note =  new SimpleNote("F#", 4); break;
                case Key.T:  note =  new SimpleNote("G-", 4); break;
                case Key.D6: note =  new SimpleNote("G#", 4); break;
                case Key.Y:  note =  new SimpleNote("A-", 4); break;
                case Key.D7: note =  new SimpleNote("A#", 4); break;
                case Key.U:  note =  new SimpleNote("B-", 4); break;
                //Next Octave
                case Key.I:  note =  new SimpleNote("C-", 5); break;
                //"Rest" Note (nothing happens)
                default: break;
            }
            return note;
        }

        /// <summary>
        /// Updates the UI Elements with the last state
        /// </summary>
        /// <param name="note"></param>
        /// <param name="k"></param>
        public void UpdateUI(SimpleNote note, Key k) { tbKey.Text = GetCharFromKey(k) + ""; tbNote.Text = note.Pitch.Replace("-", " ") + note.Octave; }

        #region Form Key Handling
        /// <summary>
        /// WPF OnKeyDown Event Handler
        /// </summary>
        /// <param name="s"></param>
        /// <param name="e"></param>
        public void wpfKeyDown(object s, KeyEventArgs e) { if (!IsKeyDown(e.Key)) { SimpleNote n = ConvertKeyToNote(e.Key); midi.PlayNote(n);UpdateUI(n,e.Key); AddKeyDown(e.Key); } }
        /// <summary>
        /// WPF OnKeyUp Event Handler
        /// </summary>
        /// <param name="s"></param>
        /// <param name="e"></param>
        public void wpfKeyUp(object s, KeyEventArgs e) { if (IsKeyDown(e.Key)) { midi.StopNote(ConvertKeyToNote(e.Key)); RemoveKeyDown(e.Key); } }
        #endregion

        #region KeyDown State Management
        private void AddKeyDown(Key key) { if (!KeysDown.Contains(key)) { KeysDown.Add(key); } }
        private void RemoveKeyDown(Key key) { KeysDown.Remove(key); }
        private bool IsKeyDown(Key key) { if (KeysDown.Contains(key)) { return true; } else { return false; } }
        #endregion

        #region Decoding Key to char
        public enum MapType : uint
        {
            MAPVK_VK_TO_VSC = 0x0,
            MAPVK_VSC_TO_VK = 0x1,
            MAPVK_VK_TO_CHAR = 0x2,
            MAPVK_VSC_TO_VK_EX = 0x3,
        }

        [DllImport("user32.dll")]
        public static extern int ToUnicode(
            uint wVirtKey,
            uint wScanCode,
            byte[] lpKeyState,
            [Out, MarshalAs(UnmanagedType.LPWStr, SizeParamIndex = 4)]
            StringBuilder pwszBuff,
            int cchBuff,
            uint wFlags);

        [DllImport("user32.dll")]
        public static extern bool GetKeyboardState(byte[] lpKeyState);

        [DllImport("user32.dll")]
        public static extern uint MapVirtualKey(uint uCode, MapType uMapType);

        public static char GetCharFromKey(Key key)
        {
            char ch = ' ';

            int virtualKey = KeyInterop.VirtualKeyFromKey(key);
            byte[] keyboardState = new byte[256];
            GetKeyboardState(keyboardState);

            uint scanCode = MapVirtualKey((uint)virtualKey, MapType.MAPVK_VK_TO_VSC);
            StringBuilder stringBuilder = new StringBuilder(2);

            int result = ToUnicode((uint)virtualKey, scanCode, keyboardState, stringBuilder, stringBuilder.Capacity, 0);
            switch (result)
            {
                case -1:
                    break;
                case 0:
                    break;
                case 1:
                    {
                        ch = stringBuilder[0];
                        break;
                    }
                default:
                    {
                        ch = stringBuilder[0];
                        break;
                    }
            }
            return ch;
        }
        #endregion

        public class SimpleNote: ISimpleNote {
            public string Pitch { set; get; }
            public int Octave { set; get; }
            public byte Velocity { set; get; }
            public SimpleNote() { Pitch = "R-"; Octave = 0; Velocity = 0x7F; /*127 is max velocity*/ }
            public SimpleNote(string pitch = "R-", int octave = 0, byte velocity = 0x7F):this() { Pitch = pitch; Octave = octave; Velocity = velocity; }
        }
    }


}
