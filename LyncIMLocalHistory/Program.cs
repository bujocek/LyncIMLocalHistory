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
        static Dictionary<Microsoft.Lync.Model.Conversation.Conversation, ConversationContainer> ActiveConversations = 
            new Dictionary<Microsoft.Lync.Model.Conversation.Conversation, ConversationContainer>();

        static Self myself;

        static int nextConvId = 0;

        static string mydocpath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        static void Main(string[] args)
        {
            var client = LyncClient.GetClient();
            myself = client.Self;
            client.ConversationManager.ConversationAdded += ConversationManager_ConversationAdded;
            client.ConversationManager.ConversationRemoved += ConversationManager_ConversationRemoved;
            Console.ReadLine();
        }

        static void ConversationManager_ConversationAdded(object sender, Microsoft.Lync.Model.Conversation.ConversationManagerEventArgs e)
        {
            Console.WriteLine("Conversation added.");
            ActiveConversations.Add(e.Conversation, new ConversationContainer()
            {
                Conversation = e.Conversation,
                ConversationCreated = DateTime.Now,
                m_convId = nextConvId++
            });
            e.Conversation.ParticipantAdded += Conversation_ParticipantAdded;
            e.Conversation.ParticipantRemoved += Conversation_ParticipantRemoved;

            //foreach (Participant participant in conversation.Participants)
            //{
            //    (participant.Modalities[ModalityTypes.InstantMessage] as InstantMessageModality).InstantMessageReceived += InstantMessageModality_InstantMessageReceived;
            //    if(participant.Contact == myself.Contact)
            //        Console.WriteLine("Found you.");
            //    else
            //        Console.WriteLine("Participant " + participant.Contact.GetContactInformation(ContactInformationType.DisplayName) + "found.");
            //}
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
            String convlog = "[" + DateTime.Now + "](" + container.m_convId + ")<" + imm.Participant.Contact.GetContactInformation(ContactInformationType.DisplayName) + "> " + args.Text;
            using (StreamWriter outfile = new StreamWriter(mydocpath + @"\LyncIMHistory.txt"))
            {
                outfile.Write(convlog);
                outfile.Close();
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
