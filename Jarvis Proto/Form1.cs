using System;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.IO;
using System.Configuration;
using System.Data.SqlClient;

namespace Jarvis_Proto
{
    public partial class Form1 : Form
    {

        SpeechRecognitionEngine speechreco = new SpeechRecognitionEngine();
        SpeechSynthesizer jarvis = new SpeechSynthesizer();
        public Form1()
        {
            InitializeComponent();
            try
            {
                speechreco.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(speechreco_SpeechRecognized);
                unLoadGrammar();
                speechreco.SetInputToDefaultAudioDevice();
                speechreco.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch (Exception ex)
            {
                jarvis.SpeakAsync("Check your microphone settings" + ex.Message);
            }
        }


        private void LoadGrammar()
        {
            try
            {
                Choices texts = new Choices();
                string[] lines = File.ReadAllLines(Environment.CurrentDirectory + "\\commands.txt");
                texts.Add(lines);
                Grammar wordsList = new Grammar(new GrammarBuilder(texts));
                speechreco.LoadGrammar(wordsList);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void unLoadGrammar()
        {
            try
            {
                Choices texts = new Choices();
                string[] lines = File.ReadAllLines(Environment.CurrentDirectory + "\\Unloadcommands.txt");
                texts.Add(lines);
                Grammar wordsList = new Grammar(new GrammarBuilder(texts));
                speechreco.LoadGrammar(wordsList);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void HeyJarvis()
        {
            speechreco.UnloadAllGrammars();
            LoadGrammar();
        }
        private void UnloadJarvis()
        {
            speechreco.UnloadAllGrammars();
            unLoadGrammar();
        }
        private void speechreco_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            int ranNum;

            Random rnd = new Random();
            richTextBox1.Text = e.Result.Text;
            string speech = e.Result.Text;
            if (speech == "hello")
            {
                jarvis.Speak("hello and how are you?");
                UnloadJarvis();
            }
            if (speech == "how are you jarvis")
            {
                jarvis.Speak("i am fine sir, what would you like to do?");
                UnloadJarvis();
            }
            if (speech == "open google")
            {
                jarvis.Speak("loading google");
                System.Diagnostics.Process.Start("http://www.google.com");
                UnloadJarvis();
            }
            if (speech == "open youtube")
            {
                jarvis.Speak("loading youtube");
                System.Diagnostics.Process.Start("http://www.youtube.com");
                UnloadJarvis();
            }
            if (speech == "bye" || speech =="see you jarvis")
            {
                jarvis.Speak("Goodbye Sir, See you next time");
                this.Close();
            }
            if (speech == "hide")
            {
                jarvis.Speak("call me when you need me i'll be right here");
                this.Hide();
                UnloadJarvis();
            }
            if (speech == "say hi jarvis")
            {
                jarvis.Speak("Hello Ladies and gentlemen!");
                UnloadJarvis();
            }
            if (speech == "whats the weather")
            {
                jarvis.Speak("i'll check for you now");
                System.Diagnostics.Process.Start("http://www.weatherzone.com.au/");
                UnloadJarvis();
            }
            if (speech == "open steam")
            {
                jarvis.Speak("opening steam");
                System.Diagnostics.Process.Start(@"S:\Steam\Steam.exe");
                UnloadJarvis();

            }
            if (speech == "open notepad")
            {
                jarvis.Speak("opening");
                System.Diagnostics.Process.Start("c:\\windows\\system32\\notepad.exe");
                UnloadJarvis();

            }
            if (speech == "close notepad")
            {
                jarvis.Speak("closing notepad");
                System.Diagnostics.Process q;
                q = System.Diagnostics.Process.GetProcessesByName("notepad")[0];
                q.Kill();
                UnloadJarvis();
            }
            if (speech == "open visual studio")
            {
                jarvis.Speak("opening visual studio");
                System.Diagnostics.Process.Start(@"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Visual Studio 2017");
                UnloadJarvis();
            }
            if (speech == "who made you")
            {
                jarvis.Speak("Abner the great made me");
                UnloadJarvis();
            }
            if (speech == "how old are you")
            {
                jarvis.Speak("Not very old, i was born on April 12, 2017 ");
                UnloadJarvis();
            }
            if (speech == "open facebook")
            {
                jarvis.Speak("opening facebook");
                System.Diagnostics.Process.Start("https://www.facebook.com/");
                UnloadJarvis();
            }
            if (speech == "close window")
            {
                jarvis.Speak("Closing Window");
                SendKeys.SendWait("%{F4}");
                jarvis.Speak("Window Closed");
                UnloadJarvis();
            }
            if (speech == "show desktop")
            {
                jarvis.Speak("agreed");
                this.WindowState = FormWindowState.Minimized;
                Type typeShell = Type.GetTypeFromProgID("Shell.Application");
                object objShell = Activator.CreateInstance(typeShell);
                typeShell.InvokeMember("MinimizeAll", System.Reflection.BindingFlags.InvokeMethod, null, objShell, null);
                jarvis.Speak("here is your desktop");
                UnloadJarvis();
            }
            if (speech == "hey jarvis" || speech == "jarvis")
            {
                HeyJarvis();
                ranNum = rnd.Next(1, 5);
                if (ranNum == 1) { jarvis.Speak("Yes sir");}
                else if (ranNum == 2) { jarvis.Speak("Yes?");}
                else if (ranNum == 3) { jarvis.Speak("Yes sir. How may I help?");}
                else if (ranNum == 4) { jarvis.Speak("Yes sir. How may I be of assistance?");}
            }
            if (speech == "whats the date today")
            {
                string date;
                date = "the date is," + System.DateTime.Now.ToString("dd MMMM", new System.Globalization.CultureInfo("en-US"));
                jarvis.SpeakAsync(date);
                date = "" + System.DateTime.Today.ToString(" yyyy");
                jarvis.Speak(date);
                UnloadJarvis();
            }
            if (speech == "what day is it")
            {
                string dayis;
                dayis = "Today is," + System.DateTime.Now.ToString("dddd", new System.Globalization.CultureInfo("en-US"));
                jarvis.SpeakAsync(dayis);
                UnloadJarvis();
            }
            if (speech == "what time is it" || speech == "whats the time")
            {
                System.DateTime now = System.DateTime.Now;
                ranNum = rnd.Next(1, 3);
                string time = now.GetDateTimeFormats('t')[0];
                if (ranNum == 1) { jarvis.Speak("Time check " + time);}
                else if (ranNum == 2) { jarvis.Speak("The time is" + time);}
                else if (ranNum == 3) { jarvis.Speak("Wait I am reading the time " + " The time is " + time);}
                UnloadJarvis();
            }
            if (speech == "show yourself" || speech == "expand")
            {
                jarvis.Speak("Expanding");
                Form1 frm = new Form1();
                this.Show();
                frm.BringToFront();
                UnloadJarvis();

            }
            if (speech == "open email")
            {
                jarvis.Speak("loading email");
                System.Diagnostics.Process.Start("https://login.yahoo.com/config/mail?&.src=ym&.intl=au");
                UnloadJarvis();
            }
            if (speech == "find jobs")
            {
                jarvis.Speak("i'll search for you");
                System.Diagnostics.Process.Start("https://www.seek.com.au/?tracking=SEM-GGL-SRC-PaidSearchAU-3698&gclid=Cj0KEQjww7zHBRCToPSj_c_WjZIBEiQAj8il5O3TUIwVOq_LFns6msF0USeDTNqGyS5JncCDWJxG1pIaAvJy8P8HAQ&gclsrc=aw.ds&dclid=CJ7e-uutodMCFYJ1vQodWhsNcA");
                jarvis.Speak("Search desired job here");
                UnloadJarvis();
            }
            if (speech == "open coin spot")
            {
                jarvis.Speak("Okay, opening coinspot");
                System.Diagnostics.Process.Start("https://www.coinspot.com.au");
                jarvis.Speak("Here's your crypto currency status sir");
                UnloadJarvis();
            }

            jarvis.SpeakCompleted += new EventHandler<SpeakCompletedEventArgs>(jarvis_SpeakCompleted);
        }

        bool listening = false;
        private void jarvis_SpeakCompleted(object sender, SpeakCompletedEventArgs e)
        {
            listening = false;
            if (!listening) return;
        }
    }
}
