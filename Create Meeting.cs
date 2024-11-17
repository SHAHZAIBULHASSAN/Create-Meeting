using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Threading.Tasks;

namespace DMT.Plugin
{
    public class ScheduleMeeting : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);
            tracingService.Trace("Plugin: Begin Execute Depth ={0}", context.Depth.ToString());
           if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity targetEntity)
            {
               
                try
                {
                    var taskEntity = service.Retrieve(targetEntity.LogicalName, targetEntity.Id, new ColumnSet(true));
                    tracingService.Trace("Retrieved sl_activities entity with ID: {0}", targetEntity.Id);

                    if (context.MessageName.ToLower() == Task.Update && taskEntity.Contains(Task.ScheduleMeeting) && Convert.ToBoolean(taskEntity[Task.ScheduleMeeting])==false)
                    {
                        List<EntityReference> Attendees = new List<EntityReference>();
                        var isConfidentialFalse = taskEntity.Contains(Task.Confidential) && taskEntity.GetAttributeValue<bool>(Task.Confidential);
                        var scheduleMeeting = taskEntity.Contains(Task.ChairmanAction) && taskEntity.GetAttributeValue<OptionSetValue>(Task.ChairmanAction).Value == 2;
                        int duration =taskEntity.GetAttributeValue<int>(Task.Duration);
                        tracingService.Trace(duration.ToString());

                        if (scheduleMeeting)
                        {
                          var meetingTitle = taskEntity.GetAttributeValue<string>(Task.Title);

                            DateTime startDate = DateTime.MinValue; // Initialize with a default value
                            DateTime endDate = DateTime.MinValue; // Initialize with a default value
                           if (targetEntity.Contains(Task.MeetingStartDate) && targetEntity[Task.MeetingStartDate] is DateTime)
                            {
                                startDate = (DateTime)targetEntity[Task.MeetingStartDate];
                                tracingService.Trace("Start Date: {0}", startDate);
                            }
                            else
                            {
                                tracingService.Trace("Start date is not available in the Post-Image.");
                            }

                            if (targetEntity.Contains(Task.Duration) && targetEntity[Task.Duration] is int)
                            {
                                endDate = startDate.AddMinutes(duration);
                                tracingService.Trace("End Date: {0}", endDate);
                            }
                            else
                            {
                                tracingService.Trace("End date is not available in the Post-Image.");
                            }

                            if (startDate != DateTime.MinValue && endDate != DateTime.MinValue)
                            {
                                tracingService.Trace("User dates: Start Date == {0} and End Date == {1}", startDate.ToString(), endDate.ToString());
                            }
                            else
                            {
                                tracingService.Trace("Unable to trace user dates due to missing data.");
                            }
                             // Initialize with a default value
                            
                            DateTime localstart= DateTime.MinValue;
                            DateTime localend = DateTime.MinValue;

                            // Retrieve the user's time zone
                            Entity userSettings = GetUserSettings(service, context.UserId);
                            if (userSettings != null && userSettings.Contains("timezonecode"))
                            {
                                int timeZoneCode = (int)userSettings["timezonecode"];
                                TimeZoneInfo userTimeZone = GetTimeZoneInfo(service, timeZoneCode);

                                if (userTimeZone != null)
                                {
                                     localstart = TimeZoneInfo.ConvertTimeFromUtc(startDate, userTimeZone);
                                     localend = TimeZoneInfo.ConvertTimeFromUtc(endDate, userTimeZone);
                                    // Do something with localDateTime
                                }
                            }

                            if (!isConfidentialFalse)
                            {
                                var fetchXml = User.FetchXmlUser;
                                var usersCollection = service.RetrieveMultiple(new FetchExpression(fetchXml));
                                usersCollection.Entities.ToList().ForEach(user => Attendees.Add(new EntityReference(User.UserLogicalName, user.Id)));


                            }
                            else
                            {
                                tracingService.Trace("Only ChairmanAction condition met. Adding specific attendees...");
                                var fetchXmlUser = User.FetchxmlUser;
                                EntityCollection user = service.RetrieveMultiple(new FetchExpression(fetchXmlUser));
                                foreach (var entity in user.Entities)
                                {
                                    Attendees.Add(new EntityReference(entity.LogicalName, entity.Id));
                                }
                                tracingService.Trace("Total Attendees added: {0}", Attendees.Count);
                            }

                            AddAttendees(taskEntity, Task.CreatedBy, Attendees, tracingService);
                            AddAttendees(taskEntity, Task.MainAssignee, Attendees, tracingService);

                            var meetingId = CreateMeeting(service, taskEntity, meetingTitle, localstart, localend, Attendees, tracingService);
                            targetEntity[Task.ScheduleMeeting] = true;
                            service.Update(targetEntity);
                            tracingService.Trace("Task entity updated successfully.");
                            tracingService.Trace($"Meeting created with ID: {meetingId} for {Attendees.Count} users.");


                        }
                        

                    }
                    else
                    {
                        if (context.MessageName.ToLower() == Task.Update)
                        {
                            var forwardto = taskEntity.Contains(Task.ChairmanAction) && taskEntity.GetAttributeValue<OptionSetValue>(Task.ChairmanAction).Value == 1;
                            if (forwardto)
                            {


                                targetEntity[Task.ScheduleMeeting] = false;
                                service.Update(targetEntity);
                            }
                        }
                       
                    }
                }
                catch (Exception ex)
                {
                    tracingService.Trace("CreateandUpdateTaskManagementHistory Plugin: {0}", ex.ToString());
                    throw;
                }
            }

            tracingService.Trace("Plugin: End Execute");
        }

        private TimeZoneInfo GetTimeZoneInfo(IOrganizationService service, int timeZoneCode)
        {
            QueryExpression query = new QueryExpression("timezonedefinition")
            {
                ColumnSet = new ColumnSet("standardname"),
                Criteria = new FilterExpression()
                {
                    Conditions =
                {
                    new ConditionExpression("timezonecode", ConditionOperator.Equal, timeZoneCode)
                }
                }
            };

            EntityCollection results = service.RetrieveMultiple(query);
            if (results.Entities.Count > 0)
            {
                string timeZoneName = results.Entities[0].GetAttributeValue<string>("standardname");
                return TimeZoneInfo.FindSystemTimeZoneById(timeZoneName);
            }

            return null;
        }
    

        private Entity GetUserSettings(IOrganizationService service, Guid userId)
        {
            QueryExpression query = new QueryExpression("usersettings")
            {
                ColumnSet = new ColumnSet("timezonecode"),
                Criteria = new FilterExpression()
                {
                    Conditions =
                {
                    new ConditionExpression("systemuserid", ConditionOperator.Equal, userId)
                }
                }
            };

            EntityCollection results = service.RetrieveMultiple(query);
            if (results.Entities.Count > 0)
            {
                return results.Entities[0];
            }

            return null;
        }

        private void AddAttendees(Entity taskEntity, string attributeName, List<EntityReference> attendees, ITracingService tracingService)
        {
            if (taskEntity.Contains(attributeName) && taskEntity[attributeName] is EntityReference attendee)
                attendees.Add(attendee);
            else
                tracingService.Trace($"{attributeName} is null or not an EntityReference.");
        }

        private Guid CreateMeeting(IOrganizationService service, Entity task, string title, DateTime start, DateTime end, List<EntityReference> attendees, ITracingService tracingService)
        {
            var meeting = new Entity(Meeting.MeetingLogicalName);
            meeting[Meeting.Subject] = title;
            tracingService.Trace("start=={0} and End ={1}", start.ToString(), end.ToString());
            meeting[Meeting.StartDate] = TimeZone.CurrentTimeZone.ToLocalTime(start);
            meeting[Meeting.EndDate] = TimeZone.CurrentTimeZone.ToLocalTime(end);
            tracingService.Trace("Local Start=={0} and End ={1}", meeting[Meeting.StartDate].ToString(), meeting[Meeting.EndDate].ToString());
            meeting[Meeting.Regarding] = new EntityReference(Task.TaskLogicalName, task.Id);

            if (attendees.Count > 0)
            {
                var attendeeCollection = new EntityCollection();
                attendees.ForEach(attendee =>
                {
                    if (attendee != null)
                    {
                        var activityParty = new Entity(ActivityParty.ActivityLogicalName);
                        activityParty[ActivityParty.PartyId] = attendee;
                        activityParty[ActivityParty.Participationtype] = new OptionSetValue(5); // Assuming 5 is for required attendee
                        activityParty[ActivityParty.Partyobjecttypecode] = new OptionSetValue(GetObjectTypeCode(attendee.LogicalName));
                        attendeeCollection.Entities.Add(activityParty);
                    }
                    else
                    {
                        tracingService.Trace("Null reference encountered in attendees collection.");
                    }
                });
                meeting[Meeting.Attendees] = attendeeCollection;
            }

            return service.Create(meeting);
        }

        private int GetObjectTypeCode(string logicalName)
        {
            switch (logicalName)
            {
                case "systemuser":
                    return 8;
                case "contact":
                    return 2;
                case "account":
                    return 1;
                default:
                    throw new InvalidOperationException("Unsupported entity type");
            }
        }
    }
}
