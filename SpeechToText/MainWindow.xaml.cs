using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Speech.Recognition;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Controls;
using System.Diagnostics;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.IO.Ports;
using System.Runtime.InteropServices;

namespace SpeechToText
{
        public partial class MainWindow : Window
        {
            SpeechRecognitionEngine speechRecognitionEngine = null;
            List<Word> words = new List<Word>();
            public MainWindow()
            {
                InitializeComponent();
                try
                {
                    speechRecognitionEngine = createSpeechEngine("en-US");
                    speechRecognitionEngine.AudioLevelUpdated += new EventHandler<AudioLevelUpdatedEventArgs>(engine_AudioLevelUpdated);
                    speechRecognitionEngine.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(engine_SpeechRecognized);
                    loadGrammarAndCommands();
                    speechRecognitionEngine.SetInputToDefaultAudioDevice();
                    //speechRecognitionEngine.SetInputToWaveFile("untitled.wmv");
                    speechRecognitionEngine.RecognizeAsync(RecognizeMode.Multiple);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show(ex.Message, "Voice recognition failed");
                }
            }
            
            static private SerialPort port = new SerialPort("COM4",9600, Parity.None, 8, StopBits.One);
            private SpeechRecognitionEngine createSpeechEngine(string preferredCulture)
            {
                foreach (RecognizerInfo config in SpeechRecognitionEngine.InstalledRecognizers())
                {
                    if (config.Culture.ToString() == preferredCulture)
                    {
                        speechRecognitionEngine = new SpeechRecognitionEngine(config);
                        break;
                    }
                }

                if (speechRecognitionEngine == null)
                {
                    System.Windows.MessageBox.Show("The desired culture is not installed on this machine, the speech-engine will continue using "
                        + SpeechRecognitionEngine.InstalledRecognizers()[0].Culture.ToString() + " as the default culture.",
                        "Culture " + preferredCulture + " not found!");
                    speechRecognitionEngine = new SpeechRecognitionEngine(SpeechRecognitionEngine.InstalledRecognizers()[0]);
                }

                return speechRecognitionEngine;
            }

            private void loadGrammarAndCommands()
            {
                try
                {
                    Choices texts = new Choices();
                    string[] lines = File.ReadAllLines(Environment.CurrentDirectory + "\\example.txt");
                    foreach (string line in lines)
                    {
                        if (line == String.Empty) continue;

                        // split the line
                        var parts = line.Split(new char[] { '|' });

                        // add commandItem to the list for later lookup or execution
                        words.Add(new Word() { Text = parts[0], AttachedText = parts[1] });

                        // add the text to the known choices of speechengine
                        texts.Add(parts[0]);
                    }
                    Grammar wordsList = new Grammar(new GrammarBuilder(texts));
                    speechRecognitionEngine.LoadGrammar(wordsList);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
    
      
            private string getKnownText(string command)
            {
                try
                {
                    var cmd = words.Where(c => c.Text == command).First();
                    return cmd.AttachedText;
                
                }
                catch (Exception)
                {
                    return command;
                }
            }

        
            void engine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
            {
                txtSpoken.Text += "\r" + getKnownText(e.Result.Text);
                scvText.ScrollToEnd();
                int i, j, k, flag = 0;
                
                string[] speech = new string[7] { "notepad", "computer", "calculator", "chrome", "word", "excel", "wordpad"};
                string[] execute = new string[7] { "notepad.exe", "explorer.exe", "calc.exe", "chrome.exe", "WINWORD.exe", "EXCEL.exe", "wordpad.exe" };
                string[] speech1 = new string[24] { "left", "right", "up", "down", "close", "go up", "page up","page down", "enter", "escape", "backspace", "home", "end" ,"tab","refresh","mail","face book","new tab","next tab","minimize","maximize","restore","close tab","change"};
                string[] keycode = new string[24] { "{LEFT}", "{RIGHT}", "{UP}", "{DOWN}", "%{F4}", "%{UP}", "{PGUP}", "{PGDN}","{ENTER}", "{ESC}", "{BS}", "{HOME}", "{END}","{TAB}","{F5}" ,"(mail)","(facebook)","^{t}","^{TAB}","% n","% x","% r","^{w}","%{TAB}"};
                string[] speech2 = new string[36] { "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z" ,"zero","one","two","three","four","five","six","seven","eight","nine"};
                string[] charkey = new string[36] { "{a}", "{b}", "{c}", "{d}", "{e}", "{f}", "{g}", "{h}", "{i}", "{j}", "{k}", "{l}", "{m}", "{n}", "{o}", "{p}", "{q}", "{r}", "{s}", "{t}", "{u}", "{v}", "{w}", "{x}", "{y}", "{z}" ,"{0}","{1}","{2}","{3}","{4}","{5}","{6}","{7}","{8}","{9}"};
                for (i = 0; i < 7; i++)
                {
                    if (e.Result.Text == speech[i])
                    {
                        Process.Start(execute[i]);
                        flag = 1;
                        port.Open();
                        port.Write("l");
                        port.Close();
                        break;
                    }
                    else
                    {
                        flag = 0;
                        continue;
                    }
                }
    
                if (flag == 0)
                {
                    for (j = 0; j < 24; j++)
                    {
                        if (e.Result.Text == speech1[j])
                        {
                            SendKeys.SendWait(keycode[j]);
                            flag = 1;
                            break;
                        }
                        else
                            continue;
                    }
                }
                if (flag == 0)
                {
                    for (k = 0; k < 36; k++)
                    {
                        if (e.Result.Text == speech2[k])
                        {
                            SendKeys.SendWait(charkey[k]);
                            flag = 1;
                            break;
                        }
                        else
                            continue;
                    }
                }
            }
            void engine_AudioLevelUpdated(object sender, AudioLevelUpdatedEventArgs e)
            {
                prgLevel.Value = e.AudioLevel;
            }

            private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
            {
                // unhook events
                speechRecognitionEngine.RecognizeAsyncStop();
                // clean references
                speechRecognitionEngine.Dispose();
            }

            private void Button_Click(object sender, RoutedEventArgs e)
           {
                this.Close();
           }

        }
}
