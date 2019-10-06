using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Recognition; //@dil Import from Reference (System.Speech) .NET FRAMEWORK 4.5 or above.
using System.Speech.Synthesis; //@dil Import from Reference (System.Speech) .NET FRAMEWORK 4.5 or above.
using System.Diagnostics; //@dil Import from Reference
using System.IO; //@dil Import from Reference
using System.Xml; //@dil Import from Reference
using System.Runtime.InteropServices; //@dil Import from Reference
namespace SpeakCodebyAdil
{
    public partial class Form1 : Form
    {  
        SpeechSynthesizer s = new SpeechSynthesizer();
        //@dil Default boolean values
        Boolean wake = true;
        Boolean offline = false;
        
        Choices list = new Choices();

        enum RecycleFlags : uint
        {
        //@dil Code to empty recycle bin on command.
            SHERB_NOCONFIRMATION = 0x00000001,
            SHERB_NOPROGRESSUI = 0x00000001,
            SHERB_NOSOUND = 0x00000004
        }
        [DllImport("Shell32.dll", CharSet = CharSet.Unicode)]
        static extern uint SHEmptyRecycleBin(IntPtr hwnd, string pszRootPath,
        RecycleFlags dwFlags);

        public Form1()
        {
            SpeechRecognitionEngine rec = new SpeechRecognitionEngine();
            //@dil commands file location eg. Open Paint
             list.Add(File.ReadAllLines(@"C:\input.txt"));

            Grammar gr = new Grammar(new GrammarBuilder(list));
            try
            {
            //@dil Speech Recognizer
                rec.RequestRecognizerUpdate();
                rec.LoadGrammar(gr);
                rec.SpeechRecognized += rec_SpeechRecognized;
                rec.SetInputToDefaultAudioDevice();
                rec.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch
            {
                return;
            }
            //@dil Select voice Male or Female
            s.SelectVoiceByHints(VoiceGender.Male);
            //@dil Speaks following text
            s.Speak("ADIL ENGINE ONLINE SUCCESS.");
            InitializeComponent();
        }
        
        //@dil Kill a program eg. Close paint
        public static void killProg(String s)
        {
            System.Diagnostics.Process[] procs = null;
            try
            {
                procs = Process.GetProcessesByName(s);
                Process prog = procs[0];
                if (!prog.HasExited)
                {
                    prog.Kill();
                }
            }
            catch
            {
            //@dil Handles exception and display box if program is not opened.
                MessageBox.Show("Program Already Closed.");
          }
            finally
            {
                if (procs != null)
                {
                    foreach (Process p in procs)
                    {
                        p.Dispose();
                    }
                }

            }
            procs = null;
        }

        public void say(String h)
        {
             //s.Speak(h);
              s.SpeakAsync(h);
            
        }

       //@dil Logoff windows command
        Process pross = new Process();
        public void lockpc()
        {

            System.Diagnostics.Process.Start(@"C:\WINDOWS\system32\rundll32.exe", "user32.dll,LockWorkStation");
            return;
        }
        private void rec_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            String r = e.Result.Text;
            
            //@dil NOTE YOU CAN USE SWITCH CASE INSTEAD "IF" CONDITON IF YOU LIKE.
            //@dil YOU CAN ADD BELOW MORE COMMANDS USING STATEMENTS and CONDITIONS.
            
            if (r == "online")
            {
                say("Im Listening!");
                wake = true;
                label1.Text = "Listening...";
              
            }
            if (r == "sleep")
            {
                say("Im Sleeping..");
                wake = false;
                label1.Text = "Sleeping...";
            }

            
            if (wake == true)

            {
                if (e.Result.Text == "offline")
                {
                    say("Have a Good day, Sir");
                    offline = true;
                }
                if(offline)
                {
                    
                    Environment.Exit(0);
                }
                if (r == "close mediaplayer")
                {
                   killProg("vlc");
                }
                if (r == "empty recyclebin")
                {
                    say("Are you sure Sir, You want to empty Recycle-bin, Confirm your Files before deleting.");

                    uint result = SHEmptyRecycleBin(IntPtr.Zero, null, 0);
                }
                //@dil Minimize or Maxamaize the window with command.
                if (r == "minimize")
                {
                    this.WindowState = FormWindowState.Minimized;
                 
                    say("Ok Sir.");
                }
                if (r == "maxamize")
                {
                    say("sure");
                    this.WindowState = FormWindowState.Normal;
                }
                if (r == "alex")
                {
                    say("Yes Sir,");
                }
                if (r == "open wordpad")
                {
                    
                    Process.Start(@"C:\Program Files\Windows NT\Accessories\wordpad.exe");
                    say("Wordpad is opened, Sir");
                }
                if (r == "close wordpad")
                {
                    
                    killProg("wordpad");
                    say("Wordpad is Closed, Sir");
                }
                if (r == "show command list")
                {
                    say("Command List on Screen.");
                    Process.Start(@"C:\input.txt");
                    say("Sir, Note that you can Add or Remove any Command according to your convinence.");
                }
                if (r == "which day is today")

                {
                    say("Today is");
                    say(DateTime.Today.ToString("dddd"));

                }
                if (r == "open notepad")
                {
                    Process.Start(@"C:\Windows\notepad.exe");
                    say("Notepad is Opened, Sir");
                }
                if (r == "close notepad")
                {
                    
                    killProg("notepad");
                    say("Notepad is Closed, Sir");
                }
                if (r == "logoff windows")
                {
                    say("Logging off Windows");
                    lockpc();
                }
                if (r == "thankyou")
                {
                    say("You're Welcome.");
                }
                if (r == "open drive e")
                {
                    say("Opening Drive E");
                    Process.Start(@"E:\");
                    say("Its Here, Sir");
                }
                if (r == "open drive f")
                {
                    say("Opening Drive F.");
                    Process.Start(@"F:\");
                    say("On your Screen, Sir");
                }
                if (r == "open movies folder")
                {
                    say("Opening Movies Folder");
                    Process.Start(@"F:\Movies");
                    say("List is here, Sir");
                }
                if (r == "play song")
                {
                    Process.Start(@"F:\Song\Music.mp4");
                    say("I am playing some songs from your collections.");
                }
                if (r == "what is todays date")

                {
                    say("Today is");
                    say(DateTime.Today.ToString("MM-dd-yyyy"));

                }
                if (r == "time")
                {
                    say("Now, time is");
                    DateTime now = DateTime.Now;
                    string time = now.GetDateTimeFormats('t')[0];
                    say(time);
                }
                if(r == "open controlpanel")
                {
                    Process.Start(@"control");
                    say("Controlpanel on your Screen.");
                }
                if (r == "open cmd")
                {
                    Process.Start(@"cmd.exe");
                    say("Command Prompt Terminal on your Screen.");
                }
                if (r == "close cmd")
                {

                    killProg("cmd");
                    say("Command Prompt is Closed, Sir");
                }
                if (r == "open paint")
                {
                    Process.Start(@"mspaint.exe");
                    say("Microsoft Paint on your Screen.");
                }
                if (r == "close paint")
                {

                    killProg("mspaint");
                    say("Paint is Closed, Sir");
                }
                if (r == "open visual studio")
                {
                    Process.Start(@"C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\devenv.exe");
                    say("Microsoft Visual Studio on your Screen. Here is a tip: you can put me on sleep mode.");
                }
                if (r == "close visual studio")
                {

                    killProg("devenv");
                    say("Dont forget to save your file after Working. Microsoft Visual Studio is Closed, Sir");
                }
            }       
        }
        
        //@dil Please do follow me on github and dont forget to give STAR to this project.
        //@dil MORE PROJECT is on the way.
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {
            

        }
    }
}
//CODE WITH LOVE BY MOHAMMAD ADIL @OneClickExecute
