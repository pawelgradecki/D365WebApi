"use strict";
var Odx = window.Odx || {};
Odx.WebAPI = Odx.WebAPI || {};
(function () {
    this.retrieve = function (entitysetname, id, async, successCallback, errorCallback) {
        id = id.replace("{", "").replace("}", "");
        var req = new XMLHttpRequest();
        if (async == undefined) {
            async = true;
        }

        req.open("GET", encodeURI(getWebAPIPath() + entitysetname + "(" + id + ")"), async);
        setStandardHeaders(req);
        req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200) {
                    var result = JSON.parse(this.response);
                    if (successCallback) {
                        successCallback(result);
                    }
                } else {
                    Xrm.Utility.alertDialog(this.statusText);
                    if (errorCallback) {
                        errorCallback();
                    }
                }
            }
        };
        req.send();
    };

    this.retrieveMultiple = function (queryString, maxpagesize, successCallback, errorCallback) {
        if(!maxpagesize){
            maxpagesize = 5000;
        }
        var req = new XMLHttpRequest();
        req.open("GET", encodeURI(getWebAPIPath() + queryString), false);
        setStandardHeaders(req);
        req.setRequestHeader("Prefer", "odata.include-annotations=\"*\",odata.maxpagesize=" + maxpagesize);
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200) {
                    var results = JSON.parse(this.response);
                    if (successCallback) {
                        successCallback(results.value);
                    }
                } else {
                    if (errorCallback) {
                        errorCallback(this.statusText);
                    }
                }
            }
        };
        req.send();
    };

    this.retrieveMultipleAsync = function(queryString, maxpagesize, successCallback, errorCallback){
        if(!maxpagesize){
            maxpagesize = 5000;
        }
        var req = new XMLHttpRequest();
        req.open("GET", encodeURI(getWebAPIPath() + queryString), true);
        setStandardHeaders(req);
        req.setRequestHeader("Prefer", "odata.include-annotations=\"*\",odata.maxpagesize=" + maxpagesize);
        req.onreadystatechange = function() {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200) {
                    var results = JSON.parse(this.response);
                    if (successCallback) {
                        successCallback(results.value);
                    }
                } else {
                    if (errorCallback) {
                        errorCallback(this.statusText);
                    }
                }
            }
        };
        req.send();
    };

    this.create = function (entitySetName, entity, successCallback, errorCallback) {
        var req = new XMLHttpRequest();
        req.open("POST", encodeURI(getWebAPIPath() + entitySetName + "?$select=createdon"), true);
        setStandardHeaders(req);
        req.setRequestHeader("Prefer", "return=representation");
        req.onreadystatechange = function () {
            if (this.readyState == 4 /* complete */) {
                req.onreadystatechange = null;
                if (this.status == 201) {
                    if (successCallback)
                        successCallback(JSON.parse(this.response));
                }
                else {
                    if (errorCallback)
                        errorCallback(XrmLab.WebAPI.errorHandler(this.response));
                }
            }
        };
        req.send(JSON.stringify(entity));
    };

    this.update = function (entitySetName, id, entityToUpdate, successCallback) {
        id = id.replace("{", "").replace("}", "");

        var req = new XMLHttpRequest();
        req.open("PATCH", encodeURI(getWebAPIPath() + entitySetName + "(" + id + ")"), true);
        setStandardHeaders(req);
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 204) {
                    if (successCallback) {
                        successCallback();
                    }
                } else {
                    Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send(JSON.stringify(entityToUpdate));
    };

    this.callUnboundAction = function (actionname, payload, async, successCallback) {
        var req = new XMLHttpRequest();
        req.open("POST", encodeURI(getWebAPIPath() + actionname), async);
        setStandardHeaders(req);
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status == 200 || this.status == 204) {
                    if (this.status == 200) {
                        var result = JSON.parse(this.response);
                    }

                    if (successCallback) {
                        successCallback(result);
                    }
                } else {
                    Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send(JSON.stringify(payload));
    };

    this.callBoundAction = function (entitysetname, id, actionname, payload, async, successCallback) {
        id = id.replace("{", "").replace("}", "");
        var req = new XMLHttpRequest();
        req.open("POST", encodeURI(getWebAPIPath() + entitysetname + "(" + id + ")/" + actionname), async);
        setStandardHeaders(req);
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status == 200 || this.status == 204) {
                    if (this.status == 200) {
                        var result = JSON.parse(this.response);
                    }

                    if (successCallback) {
                        successCallback(result);
                    }
                } else {
                    Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };

        if (payload) {
            req.send(JSON.stringify(payload));
        }
        else {
            req.send();
        }      
    };

    //Internal supporting functions
    function getClientUrl() {
        //Get the organization URL
        if (typeof GetGlobalContext == "function" &&
            typeof GetGlobalContext().getClientUrl == "function") {
            return GetGlobalContext().getClientUrl();
        }
        else {
            //If GetGlobalContext is not defined check for Xrm.Page.context;
            if (typeof Xrm != "undefined" &&
                typeof Xrm.Page != "undefined" &&
                typeof Xrm.Page.context != "undefined" &&
                typeof Xrm.Page.context.getClientUrl == "function") {
                try {
                    return Xrm.Page.context.getClientUrl();
                } catch (e) {
                    throw new Error("Xrm.Page.context.getClientUrl is not available.");
                }
            }
            else { throw new Error("Context is not available."); }
        }
    }

    function getWebAPIPath() {
        return getClientUrl() + "/api/data/v8.2/";
    }

    function setStandardHeaders(req) {
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    }

    // This function is called when an error callback parses the JSON response
    // It is a public function because the error callback occurs within the onreadystatechange 
    // event handler and an internal function would not be in scope.
    this.errorHandler = function (resp) {
        try {
            return JSON.parse(resp).error;
        } catch (e) {
            return new Error("Unexpected Error")
        }
    }

}).call(Odx.WebAPI);
