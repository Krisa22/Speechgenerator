using System;
using Excel = Microsoft.Office.Interop.Excel;
using Synthesizer = System.Speech.Synthesis.SpeechSynthesizer;
using System.Data;
using System.Linq;
using System.Speech.AudioFormat;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

namespace Speech2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private Synthesizer synthesizer = new Synthesizer();
        Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();
        void ExcelRead()
        {
            synthesizer.Volume = 80;
            Invoke(new Action(() =>
            {
                synthesizer.SelectVoice(listBox1.SelectedItem.ToString());
            }));

            for (int i = 1; i >= 0; i++)
            {
                Excel.Range B = ex.get_Range("B" + Convert.ToString(i), Missing.Value);
                Excel.Range C = ex.get_Range("C" + Convert.ToString(i), Missing.Value);
                if (B.Text != "")
                {
                    WAVLocationDialog.FileName = B.Text + ".wav";
                    synthesizer.SetOutputToWaveFile(WAVLocationDialog.FileName, new SpeechAudioFormatInfo(11025, AudioBitsPerSample.Eight, AudioChannel.Mono));
                    synthesizer.Speak(C.Text);
                }
                else
                {
                    ex.Workbooks.Close();
                    Environment.Exit(0);
                }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            foreach (var v in synthesizer.GetInstalledVoices().Select(v => v.VoiceInfo))
            {
                listBox1.Items.Add(v.Description);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                Excel.ShowDialog();
                ex.Workbooks.Open(Excel.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ex.Sheets[1];
                textBox1.Text = Excel.FileName.ToString();
            }
            catch
            {

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Thread thr = new Thread(ExcelRead);
            thr.Start();
        }
    }
}
