using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Conversation;
using System;
using System.IO;
using System.Collections.Generic;

namespace LyncIMLocalHistory
{
    class ConversationContainer
    {
        public Microsoft.Lync.Model.Conversation.Conversation Conversation { get; set; }
        public DateTime ConversationCreated { get; set; }
        public int m_convId;
    }

    class Program
    {
        static string welcomeText =
@"
LyncIMLocalHistory ver. 1.0
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
gbl@bujok.cz
";


        static Dictionary<Microsoft.Lync.Model.Conversation.Conversation, ConversationContainer> ActiveConversations = 
            new Dictionary<Microsoft.Lync.Model.Conversation.Conversation, ConversationContainer>();

        static Self myself;

        static int nextConvId = 0;

        static string mydocpath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        static void Main(string[] args)
        {
            Console.WriteLine(welcomeText);
            Console.WriteLine("Connecting to Lync Client...");
            LyncClient client = LyncClient.GetClient();
            myself = client.Self;
            if (!Directory.Exists(mydocpath + @"\LyncIMHistory"))
                Directory.CreateDirectory(mydocpath + @"\LyncIMHistory");
            client.ConversationManager.ConversationAdded += ConversationManager_ConversationAdded;
            client.ConversationManager.ConversationRemoved += ConversationManager_ConversationRemoved;
            Console.WriteLine("Ready!");
            Console.WriteLine();
            Console.ReadLine();
        }

        static void ConversationManager_ConversationAdded(object sender, Microsoft.Lync.Model.Conversation.ConversationManagerEventArgs e)
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
            Console.WriteLine("Conversation {0} added.", newcontainer.m_convId);
        }

        static void Conversation_ParticipantRemoved(object sender, Microsoft.Lync.Model.Conversation.ParticipantCollectionChangedEventArgs args)
        {
            (args.Participant.Modalities[ModalityTypes.InstantMessage] as InstantMessageModality).InstantMessageReceived -= InstantMessageModality_InstantMessageReceived;
            if (args.Participant.Contact == myself.Contact)
                Console.WriteLine("You removed.");
            else
                Console.WriteLine("Participant removed: " + args.Participant.Contact.GetContactInformation(ContactInformationType.DisplayName));
        }

        static void Conversation_ParticipantAdded(object sender, Microsoft.Lync.Model.Conversation.ParticipantCollectionChangedEventArgs args)
        {
            (args.Participant.Modalities[ModalityTypes.InstantMessage] as InstantMessageModality).InstantMessageReceived += InstantMessageModality_InstantMessageReceived;
            if (args.Participant.Contact == myself.Contact)
                Console.WriteLine("You added.");
            else
                Console.WriteLine("Participant added: " + args.Participant.Contact.GetContactInformation(ContactInformationType.DisplayName));
        }

        static void InstantMessageModality_InstantMessageReceived(object sender, Microsoft.Lync.Model.Conversation.MessageSentEventArgs args)
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
                //Console.WriteLine(filename);
                using (StreamWriter partfile = new StreamWriter(filename, true))
                {
                    partfile.WriteLine(convlog);
                    partfile.Close();
                }
            }

            Console.WriteLine(convlog);
        }

        static void ConversationManager_ConversationRemoved(object sender, Microsoft.Lync.Model.Conversation.ConversationManagerEventArgs e)
        {
            string ConversationID = e.Conversation.Properties[ConversationProperty.Id].ToString();
            e.Conversation.ParticipantAdded -= Conversation_ParticipantAdded;
            e.Conversation.ParticipantRemoved -= Conversation_ParticipantRemoved;
            if (ActiveConversations.ContainsKey(e.Conversation))
            {
                ConversationContainer container = ActiveConversations[e.Conversation];
                TimeSpan conversationLength = DateTime.Now.Subtract(container.ConversationCreated);
                Console.WriteLine("Conversation {0} lasted {1} seconds", container.m_convId, conversationLength);
                ActiveConversations.Remove(e.Conversation);
            }
        }
       
    }
}
