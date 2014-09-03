using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Conversation;
using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace LyncIMLocalHistory
{
    class Program:System.Windows.Forms.Form
    {
        public string welcomeText =
@"LyncIMLocalHistory ver. 1.0
===========================
Simple IM conversation tracker for people who want to keep the conversation 
history and can not use lync for it directly (i.e. it may be disabled by corp).
Conversations are stored in [your documents]\LyncIMHistory folder.

LICENCE
=======
This program is a Beerware and is distributed under GBL (General Beer Licence).
Which meens when you buy me a beer you can use modify and do whatever you 
please with the program. I don't take any responsibilities whatsoever. 

contact:
Jonas Bujok
gbl@bujok.cz";

        Program()
        {

            InitializeComponent();
            this.textBox1.Text = welcomeText;
            connectAndPrepare();
        }

        static Dictionary<Microsoft.Lync.Model.Conversation.Conversation, ConversationContainer> ActiveConversations = 
            new Dictionary<Microsoft.Lync.Model.Conversation.Conversation, ConversationContainer>();

        /**
         * this user (participant) using the lync
         */
        static Self myself;

        static int nextConvId = 0;
        private TextBox textBox1;
        private TextBox consoleBox;

        static string mydocpath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        static Program ProgramRef;

        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            ProgramRef = new Program();
            Application.Run(ProgramRef);
        }

        void consoleWriteLine(String text = "")
        {
            Console.WriteLine(text);
            if (this.consoleBox.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(consoleWriteLine);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                this.consoleBox.Text += text + System.Environment.NewLine;
            }
        }

        void connectAndPrepare()
        {
            Console.WriteLine(welcomeText);
            LyncClient client = null;
            bool tryAgain = false;
            int attempts = 0;
            int waittime = 5;
            do
            {
                tryAgain = false;
                attempts++;
                try
                {
                    if(attempts > 1)
                        consoleWriteLine(String.Format("Connecting to Lync Client. Attempt {0}...", attempts));
                    else
                        consoleWriteLine("Connecting to Lync Client...");
                    client = LyncClient.GetClient();
                }
                catch (LyncClientException _exception)
                {
                    tryAgain = true;
                    if(attempts <= 20)
                    {
                        consoleWriteLine(String.Format("Client not found. Trying again in {0} seconds.", waittime));
                        System.Threading.Thread.Sleep(waittime * 1000);
                    }
                    else
                    {
                        consoleWriteLine("Client not found. Too many attempts. Giving up.");
                        Console.ReadLine();
                        return;
                    }
                }
            } while (tryAgain);
            myself = client.Self;
            if (!Directory.Exists(mydocpath + @"\LyncIMHistory"))
                Directory.CreateDirectory(mydocpath + @"\LyncIMHistory");
            client.ConversationManager.ConversationAdded += ConversationManager_ConversationAdded;
            client.ConversationManager.ConversationRemoved += ConversationManager_ConversationRemoved;
            consoleWriteLine("Ready!");
            consoleWriteLine();
            Console.ReadLine();
        }

        void ConversationManager_ConversationAdded(object sender, Microsoft.Lync.Model.Conversation.ConversationManagerEventArgs e)
        {
            ConversationContainer newcontainer = new ConversationContainer()
            {
                Conversation = e.Conversation,
                ConversationCreated = DateTime.Now,
                m_convId = nextConvId++
            };
            ActiveConversations.Add(e.Conversation, newcontainer);
            e.Conversation.ParticipantAdded += Conversation_ParticipantAdded;
            e.Conversation.ParticipantRemoved += Conversation_ParticipantRemoved;
            consoleWriteLine(String.Format("Conversation {0} added.", newcontainer.m_convId));
        }

        void Conversation_ParticipantRemoved(object sender, Microsoft.Lync.Model.Conversation.ParticipantCollectionChangedEventArgs args)
        {
            (args.Participant.Modalities[ModalityTypes.InstantMessage] as InstantMessageModality).InstantMessageReceived -= InstantMessageModality_InstantMessageReceived;
            if (args.Participant.Contact == myself.Contact)
                consoleWriteLine("You removed.");
            else
                consoleWriteLine("Participant removed: " + args.Participant.Contact.GetContactInformation(ContactInformationType.DisplayName));
        }

        void Conversation_ParticipantAdded(object sender, Microsoft.Lync.Model.Conversation.ParticipantCollectionChangedEventArgs args)
        {
            (args.Participant.Modalities[ModalityTypes.InstantMessage] as InstantMessageModality).InstantMessageReceived += InstantMessageModality_InstantMessageReceived;
            if (args.Participant.Contact == myself.Contact)
                consoleWriteLine("You added.");
            else
                consoleWriteLine("Participant added: " + args.Participant.Contact.GetContactInformation(ContactInformationType.DisplayName));
        }

        void InstantMessageModality_InstantMessageReceived(object sender, Microsoft.Lync.Model.Conversation.MessageSentEventArgs args)
        {
            InstantMessageModality imm = (sender as InstantMessageModality);
            ConversationContainer container = ActiveConversations[imm.Conversation];
            DateTime now = DateTime.Now;
            String convlog = "[" + now + "] (Conv. #" + container.m_convId + ") <" + imm.Participant.Contact.GetContactInformation(ContactInformationType.DisplayName) + ">";
            convlog += Environment.NewLine + args.Text;
            using (StreamWriter outfile = new StreamWriter(mydocpath + @"\LyncIMHistory\AllLyncIMHistory.txt", true))
            {
                outfile.WriteLine(convlog);
                outfile.Close();
            }
            foreach (Participant participant in container.Conversation.Participants)
            {
                if (participant.Contact == myself.Contact)
                    continue;
                String directory = mydocpath + @"\LyncIMHistory\" + participant.Contact.GetContactInformation(ContactInformationType.DisplayName);
                if (!Directory.Exists(directory))
                    Directory.CreateDirectory(directory);
                string dateString = now.ToString("yyyy-MM-dd");
                String filename = directory + @"\" + dateString + ".txt";
                //consoleWriteLine(filename);
                using (StreamWriter partfile = new StreamWriter(filename, true))
                {
                    partfile.WriteLine(convlog);
                    partfile.Close();
                }
            }

            consoleWriteLine(convlog);
        }

        void ConversationManager_ConversationRemoved(object sender, Microsoft.Lync.Model.Conversation.ConversationManagerEventArgs e)
        {
            string ConversationID = e.Conversation.Properties[ConversationProperty.Id].ToString();
            e.Conversation.ParticipantAdded -= Conversation_ParticipantAdded;
            e.Conversation.ParticipantRemoved -= Conversation_ParticipantRemoved;
            if (ActiveConversations.ContainsKey(e.Conversation))
            {
                ConversationContainer container = ActiveConversations[e.Conversation];
                TimeSpan conversationLength = DateTime.Now.Subtract(container.ConversationCreated);
                consoleWriteLine(String.Format("Conversation {0} lasted {1} seconds", container.m_convId, conversationLength));
                ActiveConversations.Remove(e.Conversation);
            }
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Program));
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.consoleBox = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox1.Location = new System.Drawing.Point(12, 12);
            this.textBox1.MinimumSize = new System.Drawing.Size(100, 50);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox1.Size = new System.Drawing.Size(447, 202);
            this.textBox1.TabIndex = 0;
            this.textBox1.TabStop = false;
            // 
            // consoleBox
            // 
            this.consoleBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.consoleBox.Location = new System.Drawing.Point(12, 220);
            this.consoleBox.Multiline = true;
            this.consoleBox.Name = "consoleBox";
            this.consoleBox.ReadOnly = true;
            this.consoleBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.consoleBox.Size = new System.Drawing.Size(447, 152);
            this.consoleBox.TabIndex = 1;
            this.consoleBox.TabStop = false;
            // 
            // Program
            // 
            this.ClientSize = new System.Drawing.Size(471, 384);
            this.Controls.Add(this.consoleBox);
            this.Controls.Add(this.textBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Program";
            this.ResumeLayout(false);
            this.PerformLayout();

        }
    }

    class ConversationContainer
    {
        public Microsoft.Lync.Model.Conversation.Conversation Conversation { get; set; }
        public DateTime ConversationCreated { get; set; }
        public int m_convId;
    }
}
