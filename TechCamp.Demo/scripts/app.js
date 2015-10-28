/// <reference path='ref/jquery.d.ts' />
/// <reference path='ref/VSS.d.ts' />
define(["require", "exports", "VSS/Service", "TFS/WorkItemTracking/RestClient", "TFS/WorkItemTracking/Contracts", "VSS/Authentication/Services"], function (require, exports, VssService, TfsWitApi, WitTrackingTypes, AuthenticationService) {
    var TechCamp;
    (function (TechCamp) {
        var Demo;
        (function (Demo) {
            var TaskGenerator = (function () {
                function TaskGenerator(container) {
                    this.container = container;
                    this.witClient = VssService.getCollectionClient(TfsWitApi.WorkItemTrackingHttpClient);
                    this.workItemId = VSS.getConfiguration().action.workItemId;
                    var self = this;
                    this.witClient.getRelationTypes().then(function (linkTypes) { return self.linkTypes = linkTypes; });
                }
                TaskGenerator.prototype.processPbiDescription = function () {
                    var self = this;
                    self.tasks = new Array();
                    var processTaskLines = function (wit) {
                        var internalTasks = new Array();
                        var description = wit.fields["System.Description"];
                        description = description.replace(/<p>/g, "").replace(/<\/p>/g, "");
                        var taskLines = description.split("<br>");
                        taskLines.forEach(function (taskLine) {
                            if (taskLine.indexOf("Task") > -1) {
                                var title = taskLine.split("&quot;")[1];
                                internalTasks.push({ title: title, parentId: wit.id });
                            }
                        });
                        var processRelations = function (witRelations) {
                            if (witRelations.relations == null) {
                                return witRelations;
                            }
                            var workItemIds = new Array();
                            for (var i = 0; i < witRelations.relations.length; i++) {
                                var tempId = witRelations.relations[i].url.match("[^/]+$");
                                if (self.getLinkTypeName(witRelations.relations[i].rel) === "Child") {
                                    var relatedWorkItemId = tempId[0];
                                    workItemIds.push(relatedWorkItemId);
                                }
                            }
                            var filterTasks = function (wits) {
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
                            };
                            if (workItemIds.length > 0) {
                                self.witClient.getWorkItems(workItemIds, ["System.Id", "System.Title"])
                                    .then(filterTasks)
                                    .then(self.addTaskToPbi.bind(self))
                                    .then(self.displaySummary.bind(self));
                            }
                            else {
                                for (var m = 0; m < internalTasks.length; m++) {
                                    self.tasks.push(internalTasks[m]);
                                }
                                self.addTaskToPbi();
                                self.displaySummary();
                            }
                            return witRelations;
                        };
                        self.witClient.getWorkItem(self.workItemId, null, null, WitTrackingTypes.WorkItemExpand.Relations).then(processRelations);
                        return wit;
                    };
                    self.witClient.getWorkItem(self.workItemId, ["System.Id", "System.Title", "System.WorkItemType", "System.State", "System.AssignedTo", "System.Description"]).then(processTaskLines);
                };
                TaskGenerator.prototype.getLinkTypeName = function (rel) {
                    for (var i = 0; i < this.linkTypes.length; i++) {
                        if (rel === this.linkTypes[i].referenceName) {
                            return this.linkTypes[i].name;
                        }
                    }
                    return null;
                };
                TaskGenerator.prototype.displaySummary = function () {
                    if (this.tasks.length === 0) {
                        this.container.append("Nothing to add!");
                        return this.container;
                    }
                    for (var i = 0; i < this.tasks.length; i++) {
                        this.container.append("Added Task: " + this.tasks[i].title + "<br />");
                    }
                    return this.container;
                };
                TaskGenerator.prototype.addTaskToPbi = function () {
                    var self = this;
                    var context = VSS.getWebContext();
                    var authTokenManager = AuthenticationService.authTokenManager;
                    var getTaskJson = function (task, baseUrl) {
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
                    };
                    var apiURI = context.collection.uri + context.project.id + "/_apis/wit/workitems/$Task?api-version=1.0";
                    var getHeader = function (token) {
                        var header = authTokenManager.getAuthorizationHeader(token);
                        return header;
                    };
                    var addTask = function (header) {
                        for (var i = 0; i < self.tasks.length; i++) {
                            var task = self.tasks[i];
                            $.ajaxSetup({ headers: { 'Authorization': header } });
                            var postData = getTaskJson(task, context.collection.uri);
                            $.ajax({
                                type: 'PATCH',
                                url: apiURI,
                                contentType: 'application/json-patch+json',
                                data: JSON.stringify(postData),
                                success: function (data) {
                                    if (console)
                                        console.log('Task added successfully');
                                },
                                error: function (error) {
                                    if (console)
                                        console.log('Error ' + error.status + ': ' + error.statusText + '; url:' + apiURI);
                                }
                            });
                        }
                        return true;
                    };
                    authTokenManager.getToken().then(getHeader).then(addTask);
                };
                return TaskGenerator;
            })();
            Demo.TaskGenerator = TaskGenerator;
        })(Demo = TechCamp.Demo || (TechCamp.Demo = {}));
    })(TechCamp || (TechCamp = {}));
    var container = $("#content");
    exports.taskGenerator = new TechCamp.Demo.TaskGenerator(container);
});
