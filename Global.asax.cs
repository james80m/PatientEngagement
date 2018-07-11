using Autofac;
using System.Web.Http;
using System.Configuration;
using System.Reflection;
using Microsoft.Bot.Builder.Azure;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Dialogs.Internals;
using Microsoft.Bot.Connector;

namespace PatientEngagementBot
{
    public class WebApiApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            // Bot Storage: This is a great spot to register the private state storage for your bot. 
            // We provide adapters for Azure Table, CosmosDb, SQL Azure, or you can implement your own!
            // For samples and documentation, see: https://github.com/Microsoft/BotBuilder-Azure

            var store = new InMemoryDataStore();

            Conversation.UpdateContainer(
                builder =>
                {
                    builder.RegisterModule(new AzureModule(Assembly.GetExecutingAssembly()));


                    builder.Register(c => store)
                         .Keyed<IBotDataStore<BotData>>(AzureModule.Key_DataStore)
                         .AsSelf()
                         .SingleInstance();

                    builder.Register(c => new CachingBotDataStore(store,
                               CachingBotDataStoreConsistencyPolicy
                               .ETagBasedConsistency))
                               .As<IBotDataStore<BotData>>()
                               .AsSelf()
                               .InstancePerLifetimeScope();
                });
            GlobalConfiguration.Configure(WebApiConfig.Register);
        }
    }
}
