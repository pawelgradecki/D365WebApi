"use strict";
var Odx = window.Odx || {};
Odx.WebAPI = Odx.WebAPI || {};
(function () {
    var self = this;
    var settings = {
        maxPageSize: 5000
    };

    this.init = function (options) {
        settings = extend(settings, options);
    };

    self.performXhrRequest = function (options) {
        /// <summary>
        /// Perform XMLHttpRequest
        /// </summary>
        /// <param name="options" type="Object">Parameters of request</param>
        var s = extend({
            method: "GET",
            url: getWebAPIPath(),
            async: true,
            headers: [],
            payload: null,
            parseResponseAsJson: true,
            successMessagePredicate: function (xhr) {
                return xhr.status === 200
            },
            successCallback: function () { },
            errorCallback: function (error) { console.log(error) }
        }, options);


        var req = new XMLHttpRequest();
        req.open(s.method, s.url, s.async);
        setOdataHeaders(req);
        s.headers.forEach(function (header, index) {
            req.setRequestHeader(header.key, header.value);
        });

        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (s.successMessagePredicate(this)) {
                    var result = this;
                    if (s.parseResponseAsJson) {
                        result = JSON.parse(this.response);
                    }

                    if (s.successCallback) {
                        s.successCallback(result);
                    }
                } else {
                    var errorMessage = getErrorMessage(this);
                    console.log(errorMessage);
                    if (s.errorCallback) {
                        s.errorCallback(errorMessage);
                    }
                }
            }
        }

        if (s.payload == null) {
            req.send();
        }
        else {
            req.send(JSON.stringify(s.payload));
        }
    }

    self.retrieve = function (entitySetName, id, successCallback, errorCallback, async = true) {
        /// <summary>
        /// Retrieves entity by name and type
        /// </summary>
        /// <param name="entitySetName" type="String">Name of the entity set</param>
        /// <param name="id" type="String">Guid id of entity to retrieve</param>
        /// <param name="successCallback" type="Function">Callback for success</param>
        /// <param name="errorCallback" type="Function">Callback for error</param>
        /// <param name="async = true"></param>
        id = trimBraces(id);
        self.performXhrRequest({
            method: "GET",
            url: getWebAPIPath() + entitySetName + "(" + id + ")",
            async: async,
            headers: [{ key: 'Prefer', value: 'odata.include-annotations="*"' }],
            parseResponseAsJson: true,
            successCallback: successCallback,
            errorCallback: errorCallback
        });
    };

    self.retrieveMultiple = function (queryString, successCallback, errorCallback, async = true) {
        /// <summary>
        /// Retrieve multiple records using query string
        /// </summary>
        /// <param name="queryString" type="String">Query string to retrieve records</param>
        /// <param name="successCallback" type="Function">Callback for success</param>
        /// <param name="errorCallback" type="Function">Callback for error</param>
        /// <param name="async = true"></param>
        self.performXhrRequest({
            method: "GET",
            url: getWebAPIPath() + queryString,
            async: async,
            headers: [{ key: 'Prefer', value: 'odata.include-annotations="*",odata.maxpagesize=' + settings.maxPageSize }],
            parseResponseAsJson: true,
            successCallback: function (result) {
                var data = result.value;
                if (successCallback) {
                    successCallback(data);
                }
            },
            errorCallback: errorCallback
        });
    };

    self.create = function (entitySetName, entityData, successCallback, errorCallback, async = true) {
        /// <summary>
        /// Create entity
        /// </summary>
        /// <param name="entitySetName" type="String">Name of the entity set</param>
        /// <param name="entityData" type="Object">JSON data of entity to create</param>
        /// <param name="successCallback" type="Function">Callback for success</param>
        /// <param name="errorCallback" type="Function">Callback for error</param>
        /// <param name="async = true"></param>
        self.performXhrRequest({
            method: "POST",
            url: getWebAPIPath() + entitySetName,
            async: async,
            parseResponseAsJson: false,
            payload: entityData,
            successMessagePredicate: function (xhr) {
                return xhr.status === 204
            },
            successCallback: function (xhr) {
                var uri = xhr.getResponseHeader("OData-EntityId");
                var regExp = /\(([^)]+)\)/;
                var matches = regExp.exec(uri);
                var newEntityId = matches[1];

                if (successCallback) {
                    successCallback(newEntityId);
                }
            },
            errorCallback: errorCallback
        });
    };

    this.createAndSelect = function (entitySetName, entityData, selectQuery, successCallback, errorCallback, async = true) {
        /// <summary>
        /// Creates record and returns it from system using given selectQuery
        /// </summary>
        /// <param name="entitySetName" type="String">Name of the entity set</param>
        /// <param name="entityData" type="Object">JSON data of entity to create</param>
        /// <param name="selectQuery" type="String">Query to perform to get entity after creation</param>
        /// <param name="successCallback" type="Function">Callback for success</param>
        /// <param name="errorCallback" type="Function">Callback for error</param>
        /// <param name="async = true"></param>
        if (selectQuery == undefined) {
            selectQuery = "$select=createdon";
        }

        self.performXhrRequest({
            method: "POST",
            url: getWebAPIPath() + entitySetName + "?" + selectQuery,
            async: async,
            parseResponseAsJson: true,
            payload: entityData,
            successMessagePredicate: function (xhr) {
                return xhr.status === 201
            },
            successCallback: successCallback,
            errorCallback: errorCallback
        });
    };

    this.update = function (entitySetName, id, entityData, successCallback, errorCallback, async = true) {
        /// <summary>
        /// Updates entity with given data
        /// </summary>
        /// <param name="entitySetName" type="String">Name of the entity set</param>
        /// <param name="id" type="String">Guid id of the updated entity</param>
        /// <param name="entityData" type="Object">JSON data to update</param>
        /// <param name="successCallback" type="Function">Callback for success</param>
        /// <param name="errorCallback" type="Function">Callback for error</param>
        /// <param name="async = true"></param>
        id = trimBraces(id);
        self.performXhrRequest({
            method: "PATCH",
            url: getWebAPIPath() + entitySetName + "(" + id + ")",
            async: async,
            payload: entityData,
            parseResponseAsJson: false,
            successMessagePredicate: function (xhr) {
                return xhr.status === 204
            },
            successCallback: successCallback,
            errorCallback: errorCallback
        });
    };

    this.delete = function (entitySetName, id, successCallback, errorCallback, async = true) {
        /// <summary>
        /// Deletes entity from system
        /// </summary>
        /// <param name="entitySetName" type="String">Name of the entity set</param>
        /// <param name="id" type="String">Guid id of the deleted entity</param>
        /// <param name="successCallback" type="Function">Callback for success</param>
        /// <param name="errorCallback" type="Function">Callback for error</param>
        /// <param name="async = true"></param>
        id = trimBraces(id);
        self.performXhrRequest({
            method: "DELETE",
            url: getWebAPIPath() + entitySetName + "(" + id + ")",
            async: async,
            parseResponseAsJson: false,
            successMessagePredicate: function (xhr) {
                return xhr.status === 204 || xhr.status === 1223
            },
            successCallback: successCallback,
            errorCallback: errorCallback
        });
    };

    this.disassociate = function (parentEntitySetName, parentId, relationshipName, childId, successCallback, errorCallback, async = true) {
        /// <summary>
        /// Disassociate two records
        /// </summary>
        /// <param name="parentEntitySetName" type="String">Parent entity set name</param>
        /// <param name="parentId" type="String">Guid id of the parent entity</param>
        /// <param name="relationshipName" type="String">Name of lookup relationship</param>
        /// <param name="childId" type="String">Guid id of the child entity</param>
        /// <param name="successCallback" type="Function">Callback for success</param>
        /// <param name="errorCallback" type="Function">Callback for error</param>
        /// <param name="async = true"></param>
        parentId = trimBraces(parentId);
        self.performXhrRequest({
            method: "DELETE",
            url: getWebAPIPath() + parentEntitySetName + "(" + parentId + ")/" + relationshipName + "(" + childId + ")/$ref",
            async: async,
            parseResponseAsJson: false,
            successMessagePredicate: function (xhr) {
                return xhr.status === 204 || xhr.status === 1223
            },
            successCallback: successCallback,
            errorCallback: errorCallback
        });
    };

    this.disassociateLookup = function (parentEntitySetName, parentId, relationshipName, successCallback, errorCallback, async = true) {
        /// <summary>
        /// Update lookup value with null
        /// </summary>
        /// <param name="parentEntitySetName" type="String">Parent entity set name</param>
        /// <param name="parentId" type="String">Guid id of the parent entity</param>
        /// <param name="relationshipName" type="String">Name of lookup relationship</param>
        /// <param name="successCallback" type="Function">Callback for success</param>
        /// <param name="errorCallback" type="Function">Callback for error</param>
        /// <param name="async = true"></param>
        parentId = trimBraces(parentId);
        self.performXhrRequest({
            method: "DELETE",
            url: getWebAPIPath() + parentEntitySetName + "(" + parentId + ")/" + relationshipName + "/$ref",
            async: async,
            parseResponseAsJson: false,
            successMessagePredicate: function (xhr) {
                return xhr.status === 204 || xhr.status === 1223
            },
            successCallback: successCallback,
            errorCallback: errorCallback
        });
    };

    this.callUnboundAction = function (actionName, payload, successCallback, errorCallback, async = true) {
        /// <summary>
        /// Call Action unboud (global)
        /// </summary>
        /// <param name="actionName" type="String">Name of action</param>
        /// <param name="payload" type="Object">JSON payload to send in POST</param>
        /// <param name="successCallback" type="Function">Callback for success</param>
        /// <param name="errorCallback" type="Function">Callback for error</param>
        /// <param name="async = true"></param>
        self.performXhrRequest({
            method: "POST",
            url: getWebAPIPath() + actionName,
            async: async,
            parseResponseAsJson: true,
            payload: payload,
            successCallback: successCallback,
            errorCallback: errorCallback
        });
    };

    this.callBoundAction = function (entitySetName, id, actionName, payload, successCallback, errorCallback, async = true) {
        /// <summary>
        /// Call Action bound to a record of entity
        /// </summary>
        /// <param name="entitySetName" type="String">Name of the entity set</param>
        /// <param name="id" type="String">Guid id of the bound entity</param>
        /// <param name="actionName" type="String">Name of action</param>
        /// <param name="payload" type="Object">JSON payload to send in POST</param>
        /// <param name="successCallback" type="Function">Callback for success</param>
        /// <param name="errorCallback" type="Function">Callback for error</param>
        /// <param name="async = true"></param>
        id = trimBraces(id);
        self.performXhrRequest({
            method: "POST",
            url: getWebAPIPath() + entitySetName + "(" + id + ")/" + actionName,
            async: async,
            parseResponseAsJson: true,
            payload: payload,
            successMessagePredicate: function (xhr) {
                return xhr.status === 204 || xhr.status === 200
            },
            successCallback: successCallback,
            errorCallback: errorCallback
        });
    };

    function getErrorMessage(xhr) {
        if (xhr.response) {
            var result = JSON.parse(xhr.response);
            if (result.error) {
                if (result.error.message) {
                    return result.error.message;
                }
            }
        }
        else {
            return xhr.statusText;
        }
    }

    function trimBraces(value) {
        return value.replace("{", "").replace("}", "");
    }

    function getClientUrl() {
        //Get the organization URL
        if (typeof GetGlobalContext === "function" &&
            typeof GetGlobalContext().getClientUrl === "function") {
            return GetGlobalContext().getClientUrl();
        }
        else {
            if (typeof Xrm === "undefined" && typeof parent.Xrm !== "undefined") {
                var Xrm = parent.Xrm;
            }
            //If GetGlobalContext is not defined check for Xrm.Page.context;
            if (typeof Xrm !== "undefined" &&
                typeof Xrm.Page !== "undefined" &&
                typeof Xrm.Page.context !== "undefined" &&
                typeof Xrm.Page.context.getClientUrl === "function") {
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

    function setOdataHeaders(req) {
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    }

    function extend() {
        if (arguments[0] == null) {
            arguments[0] = {};
        }

        for (var i = 1; i < arguments.length; i++) {
            for (var key in arguments[i]) {
                if (arguments[i].hasOwnProperty(key)) {
                    if (typeof arguments[0][key] === 'object'
                        && typeof arguments[i][key] === 'object') {
                        arguments[0][key] = extend(arguments[0][key], arguments[i][key]);
                    }
                    else {
                        arguments[0][key] = arguments[i][key];
                    }
                }
            }
        }

        return arguments[0];
    }
}).call(Odx.WebAPI);