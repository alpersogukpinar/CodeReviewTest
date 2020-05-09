using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using Microsoft.VisualStudio.Services.WebApi.Patch;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CodeReviewTestApp
{
    class Program
    {
        const String collectionUri = "https://dev.azure.com/******/";
        const String teamProjectName = "ProjectName";
        const string pat = "******************";

        static void Main(string[] args)
        {
            VssCredentials creds = new VssBasicCredential(string.Empty, pat);
            VssConnection connection = new VssConnection(new Uri(collectionUri), creds);

            //Console.WriteLine("******Bug List******");
            //var bugList = getBugs(connection);
            Console.WriteLine("******My Code Reviews List******");
            var codeReviewsList = getMyCodeReviews(connection);
            Console.WriteLine("******Update Code Review******");
            Console.WriteLine("Please enter CodeReviewRequest Id : ");
            int id = int.Parse(Console.ReadLine());
            ChangeFieldValue(connection, id);


        }


  
        public static WorkItem ChangeFieldValue(VssConnection connection, int id)
        {

            JsonPatchDocument patchDocument = new JsonPatchDocument();

            patchDocument.Add(
                new JsonPatchOperation()
                {
                    Operation = Operation.Add,
                    Path = "/fields/System.State",
                    Value = "Closed"
                }
            );

            patchDocument.Add(
                new JsonPatchOperation()
                {
                    Operation = Operation.Add,
                    Path = "/fields/Microsoft.VSTS.CodeReview.ClosedStatusCode",
                    Value = "2"
                }
            );

            //patchDocument.Add(
            //    new JsonPatchOperation()
            //    {
            //        Operation = Operation.Add,
            //        Path = "/fields/Microsoft.VSTS.CodeReview.ClosedStatus",
            //        Value = "Checked-in"
            //    }
            //);

            patchDocument.Add(
                new JsonPatchOperation()
                {
                    Operation = Operation.Add,
                    Path = "/fields/Microsoft.VSTS.CodeReview.ClosingComment",
                    Value = "bla bla bla"
                }
            );

            WorkItemTrackingHttpClient workItemTrackingClient = connection.GetClient<WorkItemTrackingHttpClient>();

            WorkItem result = workItemTrackingClient.UpdateWorkItemAsync(patchDocument, id).Result;

            return result;
        }

        public static List<WorkItem> getMyCodeReviews(VssConnection connection)
        {
            WorkItemTrackingHttpClient witClient = connection.GetClient<WorkItemTrackingHttpClient>();
            List<WorkItem> myCodeReviews = new List<WorkItem>();

            Wiql codeReviewsQuery = new Wiql()
            {
                Query = @"SELECT
                            [System.Id],
                            [System.Links.LinkType],
                            [System.Title],
                            [System.State],
                            [System.Reason],
                            [System.AssignedTo],
                            [Microsoft.VSTS.Common.StateCode],
                            [Microsoft.VSTS.Common.ClosedBy],
                            [Microsoft.VSTS.CodeReview.ClosedStatus],
                            [Microsoft.VSTS.CodeReview.ClosedStatusCode],
                            [Microsoft.VSTS.CodeReview.ClosingComment],
                            [System.CommentCount]
                        FROM workitemLinks
                        WHERE
                            (
                                [Source].[System.TeamProject] = @project
                                AND [Source].[System.WorkItemType] IN GROUP 'Microsoft.CodeReviewRequestCategory'
                            )
                            AND (
                                [System.Links.LinkType] = 'System.LinkTypes.Hierarchy-Forward'
                            )
                            AND (
                                [Target].[System.WorkItemType] IN GROUP 'Microsoft.CodeReviewResponseCategory'
                            )
                        ORDER BY [System.CreatedDate] DESC,
                            [System.Id]
                        MODE (Recursive, ReturnMatchingChildren)"
                //MODE (MustContain)"
            };
            WorkItemQueryResult result = witClient.QueryByWiqlAsync(codeReviewsQuery, teamProjectName).Result;


            if (result.WorkItemRelations.Any())
            {
                int skip = 0;
                const int batchSize = 100;
                IEnumerable<WorkItemLink> workItemLinks;
                
                do
                {
                    result.WorkItemRelations.Skip(skip).Take(batchSize);
                    workItemLinks = result.WorkItemRelations.Skip(skip).Take(batchSize);
                    if (workItemLinks.Any())
                    {
                        // get details for each work item in the batch
                        //List<WorkItem> workItems = witClient.GetWorkItemsAsync(workItemLinks.Select(wir => wir.Target.Id), null, null, WorkItemExpand.All).Result;
                        List<string> fields = new List<String>() {
                            "System.Title", "System.WorkItemType", 
                            "System.State", "Microsoft.VSTS.Common.StateCode", 
                            "Microsoft.VSTS.CodeReview.ClosedStatus", 
                            "Microsoft.VSTS.CodeReview.ClosedStatusCode", 
                            "Microsoft.VSTS.CodeReview.ClosingComment" };
                        List<WorkItem> workItems = witClient.GetWorkItemsAsync(workItemLinks.Select(wir => wir.Target.Id), fields, null,null).Result;

                        foreach (WorkItem workItem in workItems)
                        {
                            myCodeReviews.Add(workItem);
                            string closedStatus = workItem.Fields.ContainsKey("Microsoft.VSTS.CodeReview.ClosedStatus") ? workItem.Fields["Microsoft.VSTS.CodeReview.ClosedStatus"].ToString() : "*";
                            string closedStatusCode = workItem.Fields.ContainsKey("Microsoft.VSTS.CodeReview.ClosedStatusCode") ? workItem.Fields["Microsoft.VSTS.CodeReview.ClosedStatusCode"].ToString() : "*";
                            // write work item to console
                            string sItem = $"WorkItemId : {workItem.Id} - " +
                                $"Title : {workItem.Fields["System.Title"]} - " +
                                $"Type : {workItem.Fields["System.WorkItemType"]} - " +
                                $"State : {workItem.Fields["System.State"]} -" +
                                $"StateCode : { workItem.Fields["Microsoft.VSTS.Common.StateCode"]} - " +
                                $"ClosedStatus : {closedStatus} - " +
                                $"ClosedStatusCode : {closedStatusCode}";

                            Console.WriteLine(sItem);                           
                        }
                    }
                    skip += batchSize;
                }
                while (workItemLinks.Count() == batchSize);

            }
            else
            {
                Console.WriteLine("No CodeReviews returned from query.");
            }
            return myCodeReviews;
        }

        public static List<WorkItem>  getBugs(VssConnection connection)
        {

            // Create instance of WorkItemTrackingHttpClient using VssConnection
            WorkItemTrackingHttpClient witClient = connection.GetClient<WorkItemTrackingHttpClient>();

            List<WorkItem> bugList = new List<WorkItem>();
            // Get 2 levels of query hierarchy items
            List <QueryHierarchyItem> queryHierarchyItems = witClient.GetQueriesAsync(teamProjectName, depth: 2).Result;

            // Search for 'My Queries' folder
            QueryHierarchyItem myQueriesFolder = queryHierarchyItems.FirstOrDefault(qhi => qhi.Name.Equals("My Queries"));
            if (myQueriesFolder != null)
            {
                string queryName = "REST Sample";

                // See if our 'REST Sample' query already exists under 'My Queries' folder.
                QueryHierarchyItem newBugsQuery = null;
                if (myQueriesFolder.Children != null)
                {
                    newBugsQuery = myQueriesFolder.Children.FirstOrDefault(qhi => qhi.Name.Equals(queryName));
                }
                if (newBugsQuery == null)
                {
                    // if the 'REST Sample' query does not exist, create it.
                    newBugsQuery = new QueryHierarchyItem()
                    {
                        Name = queryName,
                        Wiql = "SELECT [System.Id],[System.WorkItemType],[System.Title],[System.AssignedTo],[System.State],[System.Tags] FROM WorkItems WHERE [System.TeamProject] = @project AND [System.WorkItemType] = 'Bug' AND [System.State] = 'New'",
                        IsFolder = false
                    };
                    newBugsQuery = witClient.CreateQueryAsync(newBugsQuery, teamProjectName, myQueriesFolder.Name).Result;
                }

                // run the 'REST Sample' query
                WorkItemQueryResult result = witClient.QueryByIdAsync(newBugsQuery.Id).Result;

                if (result.WorkItems.Any())
                {
                    int skip = 0;
                    const int batchSize = 100;
                    IEnumerable<WorkItemReference> workItemRefs;
                    do
                    {
                        workItemRefs = result.WorkItems.Skip(skip).Take(batchSize);
                        if (workItemRefs.Any())
                        {
                            // get details for each work item in the batch
                            List<WorkItem> workItems = witClient.GetWorkItemsAsync(workItemRefs.Select(wir => wir.Id)).Result;
                            foreach (WorkItem workItem in workItems)
                            {
                                bugList.Add(workItem);
                                // write work item to console
                                Console.WriteLine("{0} {1}", workItem.Id, workItem.Fields["System.Title"]);
                            }
                        }
                        skip += batchSize;
                    }
                    while (workItemRefs.Count() == batchSize);
                    
                }
                else
                {
                    Console.WriteLine("No bugs were returned from query.");                    
                }
            }

            return bugList;
        }
    }
}
