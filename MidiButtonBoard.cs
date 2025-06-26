using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace MidiButtonBoard
{
    public class MidiBoard : Form
    {
        private Button[] buttons = new Button[48]; // 48 keys (C2 to B5)
        private IntPtr midiOutHandle = IntPtr.Zero;
        private NumericUpDown sizeInput;
        private TextBox[] labelInputs = new TextBox[48];
        private NumericUpDown[] noteInputs = new NumericUpDown[48];
        private NumericUpDown[] centerVolumeInputs = new NumericUpDown[48]; // Loudest (bottom)
        private NumericUpDown[] midVolumeInputs = new NumericUpDown[48];   // Second loudest (1/3 from bottom)
        private NumericUpDown[] edgeVolumeInputs = new NumericUpDown[48]; // Quietest (2/3 from bottom)
        private int whiteKeyWidth = 30; // Base white key width
        private int whiteKeyHeight = 150; // Base white key height (5x width)
        private int blackKeyWidth; // Set in UpdateKeySizes
        private int blackKeyHeight; // Set in UpdateKeySizes
        private const int GridMargin = 10;
        private const int ControlSpacing = 5; // Spacing for right-side controls
        private readonly Color TextBoxForeColor = Color.FromArgb(61, 171, 22);
        private bool[] isButtonProcessing = new bool[48];
        private Timer[] noteOffTimers = new Timer[48];
        private readonly int[] keyPattern = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11 }; // C, C#, D, D#, E, F, F#, G, G#, A, A#, B
        private readonly bool[] isBlackKey = { false, true, false, true, false, false, true, false, true, false, true, false }; // Black key flags

        [DllImport("winmm.dll")]
        private static extern int midiOutOpen(out IntPtr handle, int deviceID, IntPtr callback, IntPtr instance, int flags);

        [DllImport("winmm.dll")]
        private static extern int midiOutClose(IntPtr handle);

        [DllImport("winmm.dll")]
        private static extern int midiOutShortMsg(IntPtr handle, uint msg);

        [DllImport("winmm.dll")]
        private static extern int midiOutGetNumDevs();

        [DllImport("winmm.dll")]
        private static extern int midiOutGetDevCaps(int deviceID, out MIDIOUTCAPS caps, int size);

        [StructLayout(LayoutKind.Sequential)]
        private struct MIDIOUTCAPS
        {
            public ushort wMid;
            public ushort wPid;
            public uint vDriverVersion;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string szPname;
            public ushort wTechnology;
            public ushort wVoices;
            public ushort wNotes;
            public ushort wChannelMask;
            public uint dwSupport;
        }

        public MidiBoard()
        {
            this.Text = "MIDI Piano";
            this.FormBorderStyle = FormBorderStyle.Sizable; // Make resizable
            this.MaximizeBox = true; // Enable maximize
            this.BackColor = Color.Black;
            this.Resize += new EventHandler(Form_Resize); // Handle resizing

            int midiYokeDeviceID = FindMidiYokeDevice();
            if (midiYokeDeviceID == -1)
            {
                MessageBox.Show("MIDI Yoke not found. Please install MIDI Yoke.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
                return;
            }

            try
            {
                int result = midiOutOpen(out midiOutHandle, midiYokeDeviceID, IntPtr.Zero, IntPtr.Zero, 0);
                if (result != 0)
                    throw new Exception(String.Format("Failed to open MIDI Yoke (Device ID: {0}). Error code: {1}", midiYokeDeviceID, result));
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("MIDI output failed: {0}\nEnsure MIDI Yoke is installed.", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
                return;
            }

            InitializeControls();
            UpdateKeySizes();
            LayoutControls();
            UpdateFormSize();
        }

        private int FindMidiYokeDevice()
        {
            int numDevices = midiOutGetNumDevs();
            string deviceList = "Available MIDI output devices:\n";
            int midiYokeDeviceID = -1;

            for (int i = 0; i < numDevices; i++)
            {
                MIDIOUTCAPS caps;
                int result = midiOutGetDevCaps(i, out caps, Marshal.SizeOf(typeof(MIDIOUTCAPS)));
                if (result == 0)
                {
                    deviceList += String.Format("Device {0}: {1}\n", i, caps.szPname);
                    if (caps.szPname.Contains("MIDI Yoke"))
                    {
                        midiYokeDeviceID = i;
                    }
                }
            }

            if (midiYokeDeviceID == -1)
            {
                MessageBox.Show(String.Format("MIDI Yoke not found. {0}\nPlease install MIDI Yoke.", deviceList), 
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return midiYokeDeviceID;
        }

        private void InitializeControls()
        {
            try
            {
                string[] noteNames = { "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B" };
                for (int i = 0; i < 48; i++)
                {
                    int octave = i / 12;
                    int noteIndex = keyPattern[i % 12];

                    buttons[i] = new Button();
                    buttons[i].Text = String.Format("{0}{1}", noteNames[noteIndex], octave + 2);
                    buttons[i].Tag = i;
                    buttons[i].FlatStyle = FlatStyle.Flat;
                    buttons[i].BackColor = isBlackKey[noteIndex] ? Color.Black : Color.White;
                    buttons[i].ForeColor = isBlackKey[noteIndex] ? Color.White : Color.Black;
                    buttons[i].Font = new Font("Arial", 8, FontStyle.Bold);
                    buttons[i].Click += new EventHandler(Button_Click);
                    this.Controls.Add(buttons[i]);
                    isButtonProcessing[i] = false;

                    noteOffTimers[i] = new Timer();
                    noteOffTimers[i].Interval = 100;
                    noteOffTimers[i].Tag = i;
                    noteOffTimers[i].Tick += new EventHandler(NoteOff_Tick);

                    labelInputs[i] = new TextBox();
                    labelInputs[i].Text = buttons[i].Text;
                    labelInputs[i].Width = 60;
                    labelInputs[i].Tag = i;
                    labelInputs[i].BackColor = Color.Black;
                    labelInputs[i].ForeColor = TextBoxForeColor;
                    labelInputs[i].TextChanged += new EventHandler(LabelInput_TextChanged);
                    this.Controls.Add(labelInputs[i]);

                    noteInputs[i] = new NumericUpDown();
                    noteInputs[i].Minimum = 0;
                    noteInputs[i].Maximum = 127;
                    noteInputs[i].Value = 36 + i; // C2 (36) to B5 (83)
                    noteInputs[i].Width = 60;
                    noteInputs[i].Tag = i;
                    noteInputs[i].BackColor = Color.Black;
                    noteInputs[i].ForeColor = TextBoxForeColor;
                    this.Controls.Add(noteInputs[i]);

                    centerVolumeInputs[i] = new NumericUpDown();
                    centerVolumeInputs[i].Minimum = 0;
                    centerVolumeInputs[i].Maximum = 127;
                    centerVolumeInputs[i].Value = 127;
                    centerVolumeInputs[i].Width = 60;
                    centerVolumeInputs[i].Tag = i;
                    centerVolumeInputs[i].BackColor = Color.Black;
                    centerVolumeInputs[i].ForeColor = TextBoxForeColor;
                    this.Controls.Add(centerVolumeInputs[i]);

                    midVolumeInputs[i] = new NumericUpDown();
                    midVolumeInputs[i].Minimum = 0;
                    midVolumeInputs[i].Maximum = 127;
                    midVolumeInputs[i].Value = 80;
                    midVolumeInputs[i].Width = 60;
                    midVolumeInputs[i].Tag = i;
                    midVolumeInputs[i].BackColor = Color.Black;
                    midVolumeInputs[i].ForeColor = TextBoxForeColor;
                    this.Controls.Add(midVolumeInputs[i]);

                    edgeVolumeInputs[i] = new NumericUpDown();
                    edgeVolumeInputs[i].Minimum = 0;
                    edgeVolumeInputs[i].Maximum = 127;
                    edgeVolumeInputs[i].Value = 40;
                    edgeVolumeInputs[i].Width = 60;
                    edgeVolumeInputs[i].Tag = i;
                    edgeVolumeInputs[i].BackColor = Color.Black;
                    edgeVolumeInputs[i].ForeColor = TextBoxForeColor;
                    this.Controls.Add(edgeVolumeInputs[i]);
                }

                sizeInput = new NumericUpDown();
                sizeInput.Minimum = 20;
                sizeInput.Maximum = 50;
                sizeInput.Value = whiteKeyWidth;
                sizeInput.Width = 60;
                sizeInput.BackColor = Color.Black;
                sizeInput.ForeColor = TextBoxForeColor;
                sizeInput.ValueChanged += new EventHandler(SizeInput_ValueChanged);
                this.Controls.Add(sizeInput);
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Failed to initialize controls: {0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateKeySizes()
        {
            blackKeyWidth = (int)(whiteKeyWidth * 0.67); // Black keys 2/3 width
            whiteKeyHeight = whiteKeyWidth * 5; // Maintain 5:1 height-to-width ratio
            blackKeyHeight = (int)(whiteKeyHeight * 0.67); // Black keys 2/3 height
        }

        private void LayoutControls()
        {
            // Ensure black keys are drawn over white keys
            for (int i = 0; i < 48; i++)
            {
                if (isBlackKey[keyPattern[i % 12]])
                    buttons[i].BringToFront();
            }

            // Place white keys side by side
            int xPos = GridMargin;
            int[] whiteKeyPositions = new int[48]; // Store x positions for all keys
            int whiteKeyCount = 0;

            for (int i = 0; i < 48; i++)
            {
                if (!isBlackKey[keyPattern[i % 12]])
                {
                    buttons[i].Size = new Size(whiteKeyWidth, whiteKeyHeight);
                    buttons[i].Location = new Point(xPos, GridMargin);
                    buttons[i].Visible = true;
                    whiteKeyPositions[i] = xPos;
                    xPos += whiteKeyWidth;
                    whiteKeyCount++;
                }
            }

            // Place black keys centered over white key boundaries
            for (int i = 0; i < 48; i++)
            {
                if (isBlackKey[keyPattern[i % 12]])
                {
                    // Find previous white key
                    int prevWhite = i - 1;
                    while (prevWhite >= 0 && isBlackKey[keyPattern[prevWhite % 12]]) prevWhite--;

                    // Center black key over boundary (right edge of prev white key)
                    int blackX = prevWhite >= 0 ?
                        whiteKeyPositions[prevWhite] + whiteKeyWidth - blackKeyWidth / 2 :
                        GridMargin;
                    buttons[i].Size = new Size(blackKeyWidth, blackKeyHeight);
                    buttons[i].Location = new Point(blackX, GridMargin); // Top-aligned
                    buttons[i].Visible = true;
                    whiteKeyPositions[i] = blackX;
                }
            }

            int keyboardWidth = whiteKeyCount * whiteKeyWidth + GridMargin * 2;

            sizeInput.Location = new Point(keyboardWidth + GridMargin, GridMargin);
            sizeInput.Visible = true;

            int inputY = GridMargin + sizeInput.Height + ControlSpacing;
            for (int i = 0; i < 48; i++)
            {
                int x = keyboardWidth + GridMargin;
                labelInputs[i].Location = new Point(x, inputY + i * (labelInputs[i].Height + ControlSpacing));
                labelInputs[i].Visible = true;
                x += labelInputs[i].Width + ControlSpacing;
                noteInputs[i].Location = new Point(x, inputY + i * (noteInputs[i].Height + ControlSpacing));
                noteInputs[i].Visible = true;
                x += noteInputs[i].Width + ControlSpacing;
                centerVolumeInputs[i].Location = new Point(x, inputY + i * (centerVolumeInputs[i].Height + ControlSpacing));
                centerVolumeInputs[i].Visible = true;
                x += centerVolumeInputs[i].Width + ControlSpacing;
                midVolumeInputs[i].Location = new Point(x, inputY + i * (midVolumeInputs[i].Height + ControlSpacing));
                midVolumeInputs[i].Visible = true;
                x += midVolumeInputs[i].Width + ControlSpacing;
                edgeVolumeInputs[i].Location = new Point(x, inputY + i * (edgeVolumeInputs[i].Height + ControlSpacing));
                edgeVolumeInputs[i].Visible = true;
            }
        }

        private void UpdateFormSize()
        {
            int whiteKeyCount = 0;
            for (int i = 0; i < 48; i++)
                if (!isBlackKey[keyPattern[i % 12]])
                    whiteKeyCount++;

            int keyboardWidth = whiteKeyCount * whiteKeyWidth + GridMargin * 2;
            int keyboardHeight = whiteKeyHeight + GridMargin * 2;
            int inputWidth = 5 * (labelInputs[0].Width + ControlSpacing) + GridMargin;
            int inputHeight = sizeInput.Height + ControlSpacing + 48 * (labelInputs[0].Height + ControlSpacing);
            this.ClientSize = new Size(Math.Max(keyboardWidth + inputWidth, 800), Math.Max(keyboardHeight, inputHeight + GridMargin));
        }

        private void SizeInput_ValueChanged(object sender, EventArgs e)
        {
            whiteKeyWidth = (int)sizeInput.Value;
            UpdateKeySizes();
            LayoutControls();
            UpdateFormSize();
        }

        private void Form_Resize(object sender, EventArgs e)
        {
            // Scale key sizes based on window width
            int minWidth = 800;
            if (this.ClientSize.Width < minWidth)
                this.ClientSize = new Size(minWidth, this.ClientSize.Height);

            int whiteKeyCount = 0;
            for (int i = 0; i < 48; i++)
                if (!isBlackKey[keyPattern[i % 12]])
                    whiteKeyCount++;

            int availableKeyboardWidth = this.ClientSize.Width - (5 * (labelInputs[0].Width + ControlSpacing) + GridMargin * 2);
            if (availableKeyboardWidth < whiteKeyCount * 20)
                availableKeyboardWidth = whiteKeyCount * 20;

            whiteKeyWidth = availableKeyboardWidth / whiteKeyCount;
            UpdateKeySizes();
            LayoutControls();
            UpdateFormSize();
        }

        private void LabelInput_TextChanged(object sender, EventArgs e)
        {
            TextBox txt = (TextBox)sender;
            int index = (int)txt.Tag;
            buttons[index].Text = txt.Text;
        }

        private void Button_Click(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            int index = (int)btn.Tag;

            if (isButtonProcessing[index])
                return;

            isButtonProcessing[index] = true;

            MouseEventArgs mouseArgs = e as MouseEventArgs;
            int velocity = (int)midVolumeInputs[index].Value; // Default to second loudest

            if (mouseArgs != null)
            {
                float clickY = mouseArgs.Y;
                float height = btn.Height;
                float normalizedY = clickY / height; // 0 at top, 1 at bottom

                if (normalizedY > 2f / 3f) // Bottom (y > 2/3 height)
                    velocity = (int)centerVolumeInputs[index].Value; // Loudest
                else if (normalizedY > 1f / 3f) // 1/3 from bottom (1/3 < y ≤ 2/3)
                    velocity = (int)midVolumeInputs[index].Value; // Second loudest
                else // 2/3 from bottom (y ≤ 1/3)
                    velocity = (int)edgeVolumeInputs[index].Value; // Quietest
            }

            try
            {
                int noteNumber = (int)noteInputs[index].Value;
                int channel = 0;

                uint noteOn = (uint)(0x90 | channel | (noteNumber << 8) | (velocity << 16));
                int result = midiOutShortMsg(midiOutHandle, noteOn);
                if (result != 0)
                    throw new Exception(String.Format("Failed to send Note On. Error code: {0}", result));

                noteOffTimers[index].Tag = new MidiData(noteNumber, channel);
                noteOffTimers[index].Start();

                Color feedbackColor = velocity == (int)centerVolumeInputs[index].Value ? Color.Red :     // Bottom
                                     velocity == (int)midVolumeInputs[index].Value ? Color.Yellow :   // 1/3 from bottom
                                     Color.Green;                                             // 2/3 from bottom
                btn.BackColor = feedbackColor;

                Timer colorTimer = new Timer();
                colorTimer.Interval = 100;
                colorTimer.Tick += new EventHandler(delegate(object s, EventArgs args)
                {
                    btn.BackColor = isBlackKey[keyPattern[index % 12]] ? Color.Black : Color.White;
                    colorTimer.Stop();
                    colorTimer.Dispose();
                });
                colorTimer.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("MIDI error: {0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isButtonProcessing[index] = false;
            }
        }

        private class MidiData
        {
            public int NoteNumber;
            public int Channel;

            public MidiData(int noteNumber, int channel)
            {
                NoteNumber = noteNumber;
                Channel = channel;
            }
        }

        private void NoteOff_Tick(object sender, EventArgs e)
        {
            Timer timer = (Timer)sender;
            MidiData data = timer.Tag as MidiData;
            int index = Array.IndexOf(noteOffTimers, timer);

            if (data != null && index >= 0)
            {
                try
                {
                    uint noteOff = (uint)(0x80 | data.Channel | (data.NoteNumber << 8) | (0 << 16));
                    int result = midiOutShortMsg(midiOutHandle, noteOff);
                    if (result != 0)
                        throw new Exception(String.Format("Failed to send Note Off. Error code: {0}", result));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(String.Format("MIDI error in NoteOff: {0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            timer.Stop();
            if (index >= 0)
                isButtonProcessing[index] = false;
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            if (midiOutHandle != IntPtr.Zero)
                midiOutClose(midiOutHandle);
            foreach (Timer timer in noteOffTimers)
                timer.Dispose();
            base.OnClosing(e);
        }

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.Run(new MidiBoard());
        }
    }
}
