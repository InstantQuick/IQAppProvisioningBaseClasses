using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using IQAppProvisioningBaseClasses.Events;
using Microsoft.SharePoint.Client;

namespace IQAppProvisioningBaseClasses.Provisioning
{
    public class RemoteEventRegistrationManager : ProvisioningManagerBase
    {
        public virtual void CreateEventHandlers(ClientContext clientContext, Web web,
            List<RemoteEventRegistrationCreator> remoteEventRegistrationCreators, string remoteHost)
        {
            if (remoteEventRegistrationCreators == null || remoteEventRegistrationCreators.Count == 0) return;
            Trace.TraceInformation("Attaching event handlers at web");

            var baseEndpointUrl = "https://" + remoteHost;

            foreach (var creator in remoteEventRegistrationCreators)
            {
                var handlerEndpointUrl = baseEndpointUrl + creator.EndpointUrl;
                if (string.IsNullOrEmpty(creator.ListTitle))
                {
                    if (creator.RemoteEventScope == RemoteEventScopes.Site)
                    {
                        AttachEventHandler(handlerEndpointUrl, clientContext.Site, creator.Eventname, creator.EventReceiverType,
                            clientContext);
                    }
                    else
                    {
                        AttachEventHandler(handlerEndpointUrl, web, creator.Eventname, creator.EventReceiverType,
                            clientContext);
                    }
                }
                else
                {
                    var list = web.Lists.GetByTitle(creator.ListTitle);
                    AttachEventHandler(handlerEndpointUrl, list, creator.Eventname, creator.EventReceiverType,
                        clientContext);
                }
            }

            clientContext.ExecuteQueryRetry();
        }

        private void AttachEventHandler(string handlerEndpoint, List list, string name, EventReceiverType receiverType,
            ClientContext clientContext)
        {
            clientContext.Load(list, l => l.Title, l => l.EventReceivers.Include(e => e.ReceiverName));
            clientContext.ExecuteQueryRetry();

            var handlersToDelete = new List<EventReceiverDefinition>();
            foreach (var eventReciever in list.EventReceivers)
            {
                if (eventReciever.ReceiverName == name)
                {
                    handlersToDelete.Add(eventReciever);
                }
            }
            if (handlersToDelete.Count > 0)
            {
                foreach (var eventReceiverDefinition in handlersToDelete)
                {
                    eventReceiverDefinition.DeleteObject();
                }
                clientContext.ExecuteQuery();
            }

            var eventReceiver = new EventReceiverDefinitionCreationInformation
            {
                EventType = receiverType,
                ReceiverName = name,
                ReceiverUrl = handlerEndpoint,
                SequenceNumber = 10000,
                Synchronization = EventReceiverSynchronization.DefaultSynchronization
            };

            list.EventReceivers.Add(eventReceiver);
            OnNotify(ProvisioningNotificationLevels.Verbose,
                "Attaching remote event handler to list " + list.Title + " | " + name);
        }

        private void AttachEventHandler(string handlerEndpoint, Web web, string name, EventReceiverType receiverType,
            ClientContext clientContext)
        {
            clientContext.Load(web, w => w.EventReceivers.Include(e => e.ReceiverName));
            clientContext.ExecuteQueryRetry();

            var handlersToDelete = new List<EventReceiverDefinition>();
            foreach (var eventReciever in web.EventReceivers)
            {
                if (eventReciever.ReceiverName == name)
                {
                    handlersToDelete.Add(eventReciever);
                }
            }
            if (handlersToDelete.Count > 0)
            {
                foreach (var eventReceiverDefinition in handlersToDelete)
                {
                    eventReceiverDefinition.DeleteObject();
                }
                clientContext.ExecuteQuery();
            }

            var eventReceiver = new EventReceiverDefinitionCreationInformation
            {
                EventType = receiverType,
                ReceiverName = name,
                ReceiverUrl = handlerEndpoint,
                SequenceNumber = 10000,
                Synchronization = EventReceiverSynchronization.DefaultSynchronization
            };

            web.EventReceivers.Add(eventReceiver);
            OnNotify(ProvisioningNotificationLevels.Verbose, "Attaching remote event handler to web | " + name);
        }

        private void AttachEventHandler(string handlerEndpoint, Site site, string name, EventReceiverType receiverType,
            ClientContext clientContext)
        {
            clientContext.Load(site, s => s.EventReceivers.Include(e => e.ReceiverName));
            clientContext.ExecuteQueryRetry();

            var handlersToDelete = new List<EventReceiverDefinition>();
            foreach (var eventReciever in site.EventReceivers)
            {
                if (eventReciever.ReceiverName == name)
                {
                    handlersToDelete.Add(eventReciever);
                }
            }
            if (handlersToDelete.Count > 0)
            {
                foreach (var eventReceiverDefinition in handlersToDelete)
                {
                    eventReceiverDefinition.DeleteObject();
                }
                clientContext.ExecuteQuery();
            }

            var eventReceiver = new EventReceiverDefinitionCreationInformation
            {
                EventType = receiverType,
                ReceiverName = name,
                ReceiverUrl = handlerEndpoint,
                SequenceNumber = 10000,
                Synchronization = EventReceiverSynchronization.DefaultSynchronization
            };

            site.EventReceivers.Add(eventReceiver);
            OnNotify(ProvisioningNotificationLevels.Verbose, "Attaching remote event handler to web | " + name);
        }

        public void DeleteAll(ClientContext ctx, List<RemoteEventRegistrationCreator> remoteEventRegistrationCreators)
        {
            DeleteAll(ctx, ctx.Web, remoteEventRegistrationCreators);
        }

        public void DeleteAll(ClientContext ctx, Web web, List<RemoteEventRegistrationCreator> remoteEventRegistrationCreators)
        {
            if (remoteEventRegistrationCreators == null || remoteEventRegistrationCreators.Count == 0) return;

            foreach (var creator in remoteEventRegistrationCreators)
            {
                try
                {
                    EventReceiverDefinitionCollection events;
                    if (string.IsNullOrEmpty(creator.ListTitle))
                    {
                        if (creator.RemoteEventScope == RemoteEventScopes.Site)
                        {
                            ctx.Site.EnsureProperty(s => s.EventReceivers);
                            events = ctx.Site.EventReceivers;
                        }
                        else
                        {
                            web.EnsureProperty(s => s.EventReceivers);
                            events = web.EventReceivers;
                        }
                    }
                    else
                    {
                        var list = web.Lists.GetByTitle(creator.ListTitle);
                        list.EnsureProperty(l => l.EventReceivers);
                        events = list.EventReceivers;
                    }
                    DeleteEventHandler(ctx, events, creator.Eventname);
                }
                catch
                {
                    //ignore
                }
            }
            ctx.ExecuteQueryRetry();
        }

        private static void DeleteEventHandler(ClientContext ctx, EventReceiverDefinitionCollection events,
            string eventName)
        {
            var remoteEvent =
                events.Where(e => e.ReceiverName == eventName).FirstOrDefault();
            if (remoteEvent != null)
            {
                remoteEvent.DeleteObject();
                ctx.ExecuteQuery();
            }
        }
    }
}