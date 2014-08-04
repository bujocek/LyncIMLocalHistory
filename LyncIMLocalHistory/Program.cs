using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Conversation;
using System;
using System.Collections.Generic;

namespace LyncIMLocalHistory
{
    class ConversationContainer
    {
        public Microsoft.Lync.Model.Conversation.Conversation Conversation { get; set; }
        public DateTime ConversationCreated { get; set; }
    }

    class Program
    {
        static Dictionary<String, ConversationContainer> ActiveConversations = new Dictionary<String, ConversationContainer>();

        static Self myself;

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
            string ConversationID = e.Conversation.Properties[ConversationProperty.Id].ToString();
            StoreConversation(e.Conversation, ConversationID);
        }

        private static void StoreConversation(Conversation conversation, string ConversationID)
        {
            ActiveConversations.Add(ConversationID, new ConversationContainer()
            {
                Conversation = conversation,
                ConversationCreated = DateTime.Now
            });

            foreach (Participant participant in conversation.Participants)
            {
                (participant.Modalities[ModalityTypes.InstantMessage] as InstantMessageModality).InstantMessageReceived += InstantMessageModality_InstantMessageReceived;
                if(participant.Contact == myself.Contact)
                    Console.WriteLine("Found you.");
                else
                    Console.WriteLine("Participant found.");
            }
            Console.WriteLine("Conversation stored.");
        }

        static void InstantMessageModality_InstantMessageReceived(object sender, Microsoft.Lync.Model.Conversation.MessageSentEventArgs args)
        {
            
            Console.WriteLine("Message received: " + args.Text);
        }

        static void ConversationManager_ConversationRemoved(object sender, Microsoft.Lync.Model.Conversation.ConversationManagerEventArgs e)
        {
            string ConversationID = e.Conversation.Properties[ConversationProperty.Id].ToString();
            if (ActiveConversations.ContainsKey(ConversationID))
            {
                var container = ActiveConversations[ConversationID];
                TimeSpan conversationLength = DateTime.Now.Subtract(container.ConversationCreated);
                Console.WriteLine("Conversation {0} lasted {1} seconds", ConversationID, conversationLength);
                ActiveConversations.Remove(ConversationID);
            }
        }
       
    }
}
