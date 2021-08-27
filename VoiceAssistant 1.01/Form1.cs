using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Recognition;  // kelime algılma için kütüphane
using System.Speech.Synthesis;    // tepki vermesi için
using System.Threading;
using System.Diagnostics;         // doysa uzantısı ve çalıştırma
using System.Runtime.InteropServices;  // mouse simüle etmek


namespace audioproject
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
        private const int MOUSEEVENTF_LEFTDOWN = 0x02;    // mouse input bit settings
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;
        private const int MOUSEEVENTF_WHEEL = 0x0800;


        string[] x = { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" }; // placeholder
        string[] y = { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
        string[] t = { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
        string[] z = { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
        public Form1()
        {
        InitializeComponent();
        }
        SpeechSynthesizer synt = new SpeechSynthesizer();  //cümle tanımlanması
        PromptBuilder pbuilder = new PromptBuilder();  //komut istemi 
        SpeechRecognitionEngine rengine = new SpeechRecognitionEngine();
        private void microphone_Click(object sender, EventArgs e)
        {
            microphone.Visible = false;
        }
        void Recognized(object sender, SpeechRecognizedEventArgs e) 
        {
            switch (e.Result.Text)
            {
                case "Start":
                    microphone.Visible = false;
                    break;
            }
            if (microphone.Visible == false)
            {
                if(e.Result.Text == x[0]) // e.Result.Text ile sistemin algıladığı texti işliyor
                {
                    try
                    {
                        Process.Start(y[0]);
                        pbuilder.ClearContent();  //clearcontent ile içindekini temizleme
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                if (e.Result.Text == x[1])
                {
                    try
                    {
                        Process.Start(y[1]);
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                if (e.Result.Text == x[2])
                {
                    try
                    {
                        Process.Start(y[2]);
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                if (e.Result.Text == x[3])
                {
                    try
                    {
                        Process.Start(y[3]);
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                if (e.Result.Text == x[4])
                {
                    try
                    {
                        Process.Start(y[4]);
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                if (e.Result.Text == x[5])
                {
                    try
                    {
                        Process.Start(y[5]);
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                if (e.Result.Text == x[6])
                {
                    try
                    {
                        Process.Start(y[6]);
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                if (e.Result.Text == x[7])
                {
                    try
                    {
                        Process.Start(y[7]);
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                if (e.Result.Text == x[8])
                {
                    try
                    {
                        Process.Start(y[8]);
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                if (e.Result.Text == x[9])
                {
                    try
                    {
                        Process.Start(y[9]);
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                if (e.Result.Text == x[10])
                {
                    try
                    {
                        Process.Start(y[10]);
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                if (e.Result.Text == x[11])
                {
                    try
                    {
                        Process.Start(y[11]);
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                if (e.Result.Text == x[12])
                {
                    try
                    {
                        Process.Start(y[12]);
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                if (e.Result.Text == x[13])
                {
                    try
                    {
                        Process.Start(y[13]);
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                if (e.Result.Text == x[14])
                {
                    try
                    {
                        Process.Start(y[14]);
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                if (e.Result.Text == x[15])
                {
                    try
                    {
                        Process.Start(y[15]);
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                if (e.Result.Text == x[16])
                {
                    try
                    {
                        Process.Start(y[16]);
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                if (e.Result.Text == x[17])
                {
                    try
                    {
                        Process.Start(y[17]);
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                if (e.Result.Text == x[18])
                {
                    try
                    {
                        Process.Start(y[18]);
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                if (e.Result.Text == x[19])
                {
                    try
                    {
                        Process.Start(y[19]);
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                if (e.Result.Text == x[20])
                {
                    try
                    {
                        Process.Start(y[20]);
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Okey i opened your program for you");
                        synt.Speak(pbuilder);
                    }
                    catch
                    {
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Sorry i couldn't find your file.");
                        synt.Speak(pbuilder);
                    }
                }
                    if (e.Result.Text == t[0])
                    {
                        try
                        {
                            Process.Start(z[0]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }
                    if (e.Result.Text == t[1])
                    {
                        try
                        {
                            Process.Start(z[1]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }
                    if (e.Result.Text == t[2])
                    {
                        try
                        {
                            Process.Start(z[2]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }
                    if (e.Result.Text == t[3])
                    {
                        try
                        {
                            Process.Start(z[3]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }
                    if (e.Result.Text == t[4])
                    {
                        try
                        {
                            Process.Start(z[4]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }
                    if (e.Result.Text == t[5])
                    {
                        try
                        {
                            Process.Start(z[5]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }
                    if (e.Result.Text == t[6])
                    {
                        try
                        {
                            Process.Start(z[6]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }
                    if (e.Result.Text == t[7])
                    {
                        try
                        {
                            Process.Start(z[7]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }
                    if (e.Result.Text == t[8])
                    {
                        try
                        {
                            Process.Start(z[8]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }
                    if (e.Result.Text == t[9])
                    {
                        try
                        {
                            Process.Start(z[9]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }
                    if (e.Result.Text == t[10])
                    {
                        try
                        {
                            Process.Start(z[10]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }
                    if (e.Result.Text == t[11])
                    {
                        try
                        {
                            Process.Start(z[11]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }
                    if (e.Result.Text == t[12])
                    {
                        try
                        {
                            Process.Start(z[12]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }
                    if (e.Result.Text == t[13])
                    {
                        try
                        {
                            Process.Start(z[13]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }
                    if (e.Result.Text == t[14])
                    {
                        try
                        {
                            Process.Start(z[14]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }
                    if (e.Result.Text == t[15])
                    {
                        try
                        {
                            Process.Start(z[15]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }
                    if (e.Result.Text == t[16])
                    {
                        try
                        {
                            Process.Start(z[16]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }
                    if (e.Result.Text == t[17])
                    {
                        try
                        {
                            Process.Start(z[17]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }
                    if (e.Result.Text == t[18])
                    {
                        try
                        {
                            Process.Start(z[18]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }
                    if (e.Result.Text == t[19])
                    {
                        try
                        {
                            Process.Start(z[19]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }
                    if (e.Result.Text == t[20])
                    {
                        try
                        {
                            Process.Start(z[20]);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened your site for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                    }

                    
                switch (e.Result.Text)
                {
                    case "Read commands":  //command okuma kısmı sırasıyla
                        int i = 0;
                        foreach (string member in commands)
                        {
                        pbuilder.ClearContent();
                        pbuilder.AppendText(commands[i]);
                        synt.Speak(pbuilder);
                        i++;
                        }
                        break;
                    case "Hide commands": 
                        listBox1.Visible = false;
                        break;
                    case "Show commands":
                        listBox1.Visible = true;
                        break;
                    case "Stop":
                        microphone.Visible = true;
                        break;
                    case "Open Chrome":
                        try
                        {
                            ProcessStartInfo chrome = new ProcessStartInfo();
                            chrome.FileName = "chrome.exe"; // registry kontrolü sağlıyor.İçinde chrome.exe aranıyor.
                            if(chrome.FileName != null)
                            {
                                Process.Start(chrome);
                                pbuilder.ClearContent();
                                pbuilder.AppendText("Okey i opened chrome for you");
                                synt.Speak(pbuilder);
                            }
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                        break;
                    case "Open Word":
                        try
                        {
                            ProcessStartInfo word = new ProcessStartInfo();
                            word.FileName = "WINWORD.exe";
                            Process.Start(word);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened word for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                        break;
                    case "Open PowerPoint":
                        try
                        {
                            ProcessStartInfo pwrpnt = new ProcessStartInfo();
                            pwrpnt.FileName = "POWERPNT.exe";
                            Process.Start(pwrpnt);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened power point for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                        break;
                    case "Open Steam":
                        try
                        {
                            ProcessStartInfo steam = new ProcessStartInfo();
                            steam.FileName = "steam.exe";
                            Process.Start(steam);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened steam for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                        break;
                    case "Open Microsoft Teams":
                        try
                        {
                            ProcessStartInfo teams = new ProcessStartInfo();
                            teams.FileName = "Teams.exe";
                            Process.Start(teams);
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i opened teams for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                        break;
                    case "Close Word":
                        try
                        {
                            Process pro;
                            pro = Process.GetProcessesByName("WINWORD")[Process.GetProcessesByName("WINWORD").Count() - 1];
                            pro.Kill();
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i closed word for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                        break;
                    case "Close Steam":
                        try
                        {
                            Process pro2;
                            pro2 = Process.GetProcessesByName("steam")[Process.GetProcessesByName("steam").Count() - 1];
                            pro2.Kill();
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i closed steam for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                        break;
                    case "Close PowerPoint":
                        try
                        {
                            Process pro1;
                            pro1 = Process.GetProcessesByName("POWERPNT")[Process.GetProcessesByName("POWERPNT").Count() - 1];
                            pro1.Kill();
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Okey i closed power point for you");
                            synt.Speak(pbuilder);
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                        break;
                    case "Close Chrome":
                        try 
                        {
                            Process[] chrome = Process.GetProcessesByName("chrome"); // açık olan chrome uygulamasını kapatır. 2 tane açıksa tek komutla 1 tanesi kapanır.
                            foreach(Process p in chrome )
                            {
                                p.CloseMainWindow();
                            }
                        }
                        catch
                        {
                            pbuilder.ClearContent();
                            pbuilder.AppendText("Sorry i couldn't find your file.");
                            synt.Speak(pbuilder);
                        }
                        break;
                    case "Left":
                        MoveMouse(Cursor.Position.X - 20, Cursor.Position.Y);
                        break;
                    case "Right":
                        MoveMouse(Cursor.Position.X + 20, Cursor.Position.Y);
                        break;
                    case "Up":
                        MoveMouse(Cursor.Position.X, Cursor.Position.Y - 20);
                        break;
                    case "Down":
                        MoveMouse(Cursor.Position.X, Cursor.Position.Y + 20);
                        break;
                    case "Jump left":
                        MoveMouse(Cursor.Position.X - 200, Cursor.Position.Y);
                        break;
                    case "Jump right":
                        MoveMouse(Cursor.Position.X + 200, Cursor.Position.Y);
                        break;
                    case "Jump up":
                        MoveMouse(Cursor.Position.X, Cursor.Position.Y - 200);
                        break;
                    case "Jump down":
                        MoveMouse(Cursor.Position.X, Cursor.Position.Y + 200);
                        break;
                    case "Left mouse": // left click
                        DoMouseClick();
                        break;
                    case "Double click": // double left click
                        DoMouseClick();
                        DoMouseClick();
                        break;
                    case "Right mouse":
                        DoMouseRightClick();
                        break;
                    case "Hold left mouse":
                        HoldLeftClick();
                        break;
                    case "Release left mouse":
                        ReleaseLeftClick();
                        break;
                    case "Scroll up":
                        ScrollUp();
                        break;
                    case "Scroll down":
                        ScrollDown();
                        break;
                    case "Open google":
                        System.Diagnostics.Process.Start("http://google.com/");
                        break;
                    case "Open facebook":
                        System.Diagnostics.Process.Start("http://facebook.com/");
                        break;
                    case "Open news":
                        System.Diagnostics.Process.Start("http://milliyet.com.tr/");
                        break;
                    case "Open youtube":
                        System.Diagnostics.Process.Start("http://youtube.com/");
                        break;
                    case "Open twitter":
                        System.Diagnostics.Process.Start("http://twitter.com/");
                        break;
                    case "Open instagram":
                        System.Diagnostics.Process.Start("http://instagram.com/");
                        break;
                    case "Open netflix":
                        System.Diagnostics.Process.Start("http://netflix.com/");
                        break;
                    case "Open mynet":
                        System.Diagnostics.Process.Start("http://mynet.com/");
                        break;
                    case "Open sahibinden":
                        System.Diagnostics.Process.Start("http://sahibinden.com/");
                        break;
                    case "Exit":
                        pbuilder.ClearContent();
                        pbuilder.AppendText("Goodbye");
                        synt.Speak(pbuilder);
                        Application.Exit();
                        break;
                }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            microphone.Visible = true;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        Choices list = new Choices(); // speech.recognition kütüphanesinden gelen
        string[] commands = { "Stop", "Start", "Show commands", "Hide commands", "Read commands", "Exit", "Close Chrome", "Close Word", "Close PowerPoint", "Close Steam", "Close Microsoft Teams", "Open Chrome", "Open Word", "Open PowerPoint", "Open Steam", "Open Microsoft Teams", "Left", "Right", "Up", "Down", "Jump left", "Jump right", "Jump up", "Jump down", "Left mouse", "Right mouse", "Hold left mouse", "Release left mouse", "Double click","Scroll down","Scroll up", "Open google", "Open facebook", "Open mynet", "Open youtube", "Open netflix", "Open twitter", "Open instagram", "Open news", "Open sahibinden" };
        private void Form1_Load(object sender, EventArgs e)
        {
            listBox1.Visible = false;
            list.Add(commands);
            Grammar gram = new Grammar(new GrammarBuilder(list)); //listeyi grammarbuilder'a çevirip onu da grammar yapıyoruz.
            foreach (string element in commands)
            {
                listBox1.Items.Add(element);
            }
            try
            {
                rengine.RequestRecognizerUpdate();  //rengine 40. satırda oluşturuldu. // engine'e yeni özellikler eklemek için durduruyor.
                rengine.LoadGrammar(gram);  //engine grammar' a yükleniyor
                rengine.SpeechRecognized += Recognized;  //tanımlanan kelimeleri ekliyor
                rengine.SetInputToDefaultAudioDevice(); // varsayılann dinleme cihazını bulmak için
                rengine.RecognizeAsync(RecognizeMode.Multiple); // kelimeleri birden fazla şekilde aynı anda algılaması için (multiple)
            }
            catch
            {
                return;
            }
        }
        private void show_Click(object sender, EventArgs e)
        {
            listBox1.Visible = true;
        }
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }
        private void button4_Click(object sender, EventArgs e)
        {
            listBox1.Visible = false;
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
          
        }
        int i = 0;
        private void button3_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(textBox2.Text) && !String.IsNullOrEmpty(textBox1.Text))
            {
                try
                {
                    if (i < 20)
                    {
                        x[i] = textBox2.Text;
                        y[i] = textBox1.Text;
                        listBox1.Items.Add(x[i]);
                        Choices newlist = new Choices();
                        newlist.Add(x[i]);
                        Grammar gram1 = new Grammar(new GrammarBuilder(newlist));
                        rengine.LoadGrammar(gram1); //oluşturduğumuz yeni kelimeyi motorumuz algılıyor.
                        MessageBox.Show("Successfully added your command.");
                        textBox1.Text = String.Empty;
                        textBox2.Text = String.Empty;
                        i++;
                    }
                    else
                    {
                        i = 0; // 21 tane yeni komut ekleme hakkımız var
                    }
                }
                catch
                {
                    MessageBox.Show("Please set your display language as English");
                }
            }
            else
            {
                MessageBox.Show("You must not leave any empty spaces."); // textboxlardan biri dolu biri boş veya iksi de boşsa uyarı verecek.
                textBox1.Text = String.Empty;
                textBox2.Text = String.Empty;
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog od = new OpenFileDialog();
            od.Multiselect = false; // tek dosya seçilebilir.
            if (od.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = od.FileName;
            }
        }
        private void MoveMouse(int x, int y)
        {
            this.Cursor = new Cursor(Cursor.Current.Handle);
            Cursor.Position = new Point(x, y);
        }
        public void DoMouseClick()
        {
            int X = (int)Cursor.Position.X;
            int Y = (int)Cursor.Position.Y;
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, X, Y, 0, 0); // _lefdown ile mouse tıklanmış şekilde duruyor or bağlacı ve left up sayesinde mouse tıklaması tamamlanıyor.
        }
        public void DoMouseRightClick()
        {
            int X = (int)Cursor.Position.X;
            int Y = (int)Cursor.Position.Y;
            mouse_event(MOUSEEVENTF_RIGHTDOWN | MOUSEEVENTF_RIGHTUP, X, Y, 0, 0);
        }
        public void HoldLeftClick()
        {
            int X = (int)Cursor.Position.X;
            int Y = (int)Cursor.Position.Y;
            mouse_event(MOUSEEVENTF_LEFTDOWN, X, Y, 0, 0); //_ left down sayesinde sol click basılı kalıyor
        }
        public void ReleaseLeftClick()
        {
            int X = (int)Cursor.Position.X;
            int Y = (int)Cursor.Position.Y;
            mouse_event(MOUSEEVENTF_LEFTUP, X, Y, 0, 0);
        }
        public void ScrollUp()
        {
            mouse_event(MOUSEEVENTF_WHEEL, 0, 0, 120, 0); 
        }
        public void ScrollDown()
        {
            mouse_event(MOUSEEVENTF_WHEEL, 0, 0, -120, 0);
        }
        int a = 0;
        private void button6_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(textBox3.Text) && !String.IsNullOrEmpty(textBox4.Text))
            {
                try
                {
                    if (a < 20)
                    {
                        t[a] = textBox3.Text;
                        z[a] = textBox4.Text;
                        listBox1.Items.Add(t[a]);
                        Choices newlist = new Choices();
                        newlist.Add(t[a]);
                        Grammar gram1 = new Grammar(new GrammarBuilder(newlist));
                        rengine.LoadGrammar(gram1);
                        MessageBox.Show("Successfully added your command.");
                        textBox4.Text = String.Empty;
                        textBox3.Text = String.Empty;
                        a++;
                    }
                    else
                    {
                        a = 0;
                    }
                }
                catch
                {
                    MessageBox.Show("Please set your display language as English");
                }
            }
            else
            {
                MessageBox.Show("You must not leave any empty spaces.");
                textBox4.Text = String.Empty;
                textBox3.Text = String.Empty;
            }
        }
        private void aboutUsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This program was made by:" + "\n" + "Selim Meyva(2016513048)" +"\n" + "meyvaselim@gmail.com" +"\n" + "İlyas Oğuz Tarlakazan(2017513060)" + "\n" + "ilyastarlakazan@gmail.com");
        }
        private void howToUseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Use your voice to command the computer." +  "\n" + "You can see the commands by saying 'Show commands' or clicking 'Show commands' button.");
        }
    }
}
