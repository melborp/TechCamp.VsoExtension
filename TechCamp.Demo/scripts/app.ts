/// <reference path='ref/jquery.d.ts' />
/// <reference path='ref/VSS.d.ts' />

import VssService = require("VSS/Service");
import TfsWitApi = require("TFS/WorkItemTracking/RestClient");
import WitTrackingTypes = require("TFS/WorkItemTracking/Contracts");
import AuthenticationService = require("VSS/Authentication/Services");

module TechCamp.Demo {
    export interface ICreatedTaskInfo {
        title: string;
        parentId: number;
    }

    export class TaskGenerator {
        container: JQuery;
        witClient: TfsWitApi.WorkItemTrackingHttpClient;
        workItemId: number;
        tasks: Array<ICreatedTaskInfo>;
        linkTypes : any;

        constructor(container: JQuery) {
            this.container = container;
            this.witClient = VssService.getCollectionClient(TfsWitApi.WorkItemTrackingHttpClient);
            this.workItemId = VSS.getConfiguration().action.workItemId;
            //this.tasks = new Array();
            var self = this;
            this.witClient.getRelationTypes().then((linkTypes) => self.linkTypes = linkTypes);
        }

        processPbiDescription() {
            var self = this;
            self.tasks = new Array();

            var processTaskLines = (wit: WitTrackingTypes.WorkItem) => {
                var internalTasks = new Array();
                //parse description into tasks
                var description = wit.fields["System.Description"];
                description = description.replace(/<p>/g, "").replace(/<\/p>/g, "");

                var taskLines = description.split("<br>");
                taskLines.forEach((taskLine) => {
                    if (taskLine.indexOf("Task") > -1) {
                        var title = taskLine.split("&quot;")[1];
                        internalTasks.push({ title: title, parentId: wit.id });
                    }
                });

                //Process relations to see if any task has been already added
                var processRelations = (witRelations) => {
                    if (witRelations.relations == null) {
                        return witRelations;
                    }
                    var workItemIds = new Array();

                    //Extract the work item ids from related items
                    for (var i = 0; i < witRelations.relations.length; i++) {
                        //TODO: also check that type is Task
                        var tempId = witRelations.relations[i].url.match("[^/]+$");
                        if (self.getLinkTypeName(witRelations.relations[i].rel) === "Child") {
                            var relatedWorkItemId = tempId[0];
                            workItemIds.push(relatedWorkItemId);
                        }
                    }

                    //See if any related task already has same title, dont add it then
                    var filterTasks = (wits) => {
                        for (var j = 0; j < internalTasks.length; j++) {
                            var add = true;
                            for (var k = 0; k < wits.length; k++) {
                                var title = wits[k].fields["System.Title"];

                                if (internalTasks[j].title === title) {
                                    add = false;
                                }
                            }
                            
                            if (add) {
                                self.tasks.push(internalTasks[j]);
                            }
                        }
                        return wits;
                    }
                    if (workItemIds.length > 0) {
                        //Query related items
                        self.witClient.getWorkItems(workItemIds, ["System.Id", "System.Title"])
                            .then(filterTasks)
                            .then(self.addTaskToPbi.bind(self))
                            .then(self.displaySummary.bind(self));
                    } else {
                        for (var m = 0; m < internalTasks.length;m++) {
                            self.tasks.push(internalTasks[m]);    
                        }
                        
                        self.addTaskToPbi();
                        self.displaySummary();
                    }

                    return witRelations;
                }

                //Query relations and process them
                self.witClient.getWorkItem(self.workItemId, null, null, WitTrackingTypes.WorkItemExpand.Relations).then(processRelations);

                return wit;
            }

            self.witClient.getWorkItem(
                self.workItemId,
                ["System.Id", "System.Title", "System.WorkItemType", "System.State", "System.AssignedTo", "System.Description"]
            ).then(processTaskLines);


        }

        //Returns the name of the link for the description
        getLinkTypeName(rel) {
            for (var i = 0; i < this.linkTypes.length; i++) {
                if (rel === this.linkTypes[i].referenceName) {
                    return this.linkTypes[i].name;
                }
            }
            return null;
        }

        displaySummary() {
            if (this.tasks.length === 0) {
                this.container.append("Nothing to add!");
                return this.container;
            }

            for (var i = 0; i < this.tasks.length; i++) {
                this.container.append("Added Task: " + this.tasks[i].title + "<br />");
            }
            return this.container;
        }

        addTaskToPbi() {
            var self = this;
            var context = VSS.getWebContext();
            var authTokenManager = AuthenticationService.authTokenManager;

            var getTaskJson = (task:ICreatedTaskInfo, baseUrl) => {
                return [
                    {
                        "op": "add",
                        "path": "/fields/System.Title",
                        "value": task.title
                    },
                    {
                        "op": "add",
                        "path": "/fields/Microsoft.VSTS.Scheduling.RemainingWork",
                        "value": 0
                    },
                    {
                        "op": "add",
                        "path": "/fields/System.Description",
                        "value": task.title
                    },
                    {
                        "op": "add",
                        "path": "/fields/System.History",
                        "value": "Auto generated task from PBI"
                    },
                    {
                        "op": "add",
                        "path": "/relations/-",
                        "value": {
                            "rel": "System.LinkTypes.Hierarchy-Reverse",
                            "url": baseUrl + "_apis/wit/workItems/" + task.parentId,
                            "attributes": {
                                "comment": "Auto generated from PBI"
                            }
                        }
                    }];
            }


            var apiURI = context.collection.uri + context.project.id + "/_apis/wit/workitems/$Task?api-version=1.0";
            var getHeader = (token) => {
                var header = authTokenManager.getAuthorizationHeader(token);
                return header;
            };
            var addTask = (header) => {
                for (var i = 0; i < self.tasks.length; i++) {
                    var task = self.tasks[i];

                    //var header = authTokenManager.getAuthorizationHeader(token);
                    $.ajaxSetup({ headers: { 'Authorization': header } });

                    var postData = getTaskJson(task, context.collection.uri);
                    $.ajax({
                        type: 'PATCH',
                        url: apiURI,
                        contentType: 'application/json-patch+json',
                        data: JSON.stringify(postData),
                        success: function(data) {
                            if (console) console.log('Task added successfully');
                        },
                        error: function(error) {
                            if (console) console.log('Error ' + error.status + ': ' + error.statusText + '; url:' + apiURI);
                        }
                    });

                    
                }
                return true;
            };
            authTokenManager.getToken().then(getHeader).then(addTask);
            
        }
    }

}

var container = $("#content");
exports.taskGenerator = new TechCamp.Demo.TaskGenerator(container);
