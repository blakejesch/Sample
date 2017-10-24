define(["require", "exports", "q", "VSS/Controls", "VSS/Controls/Menus", "VSS/Controls/StatusIndicator", "applicationinsights-js"], function (require, exports, Q, Controls, Menus, StatusIndicator, AppInsights) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    var SprintGoalDto = (function () {
        function SprintGoalDto() {
        }
        return SprintGoalDto;
    }());
    exports.SprintGoalDto = SprintGoalDto;
    var SprintGoal = (function () {
        function SprintGoal() {
            var _this = this;
            this.buildWaitControl = function () {
                var waitControlOptions = {
                    target: $("#sprint-goal"),
                    cancellable: false,
                    backgroundColor: "#ffffff",
                    message: "Processing your Sprint Goal..",
                    showDelay: 0
                };
                _this.waitControl = Controls.create(StatusIndicator.WaitControl, $("#sprint-goal"), waitControlOptions);
            };
            this.getLocation = function (href) {
                var l = document.createElement("a");
                l.href = href;
                return l;
            };
            this.buildMenuBar = function () {
                var menuItems = [
                    { id: "save", text: "Save", icon: "icon-save" }
                ];
                var menubarOptions = {
                    items: menuItems,
                    executeAction: function (args) {
                        var command = args.get_commandName();
                        switch (command) {
                            case "save":
                                _this.saveSettings().then(function () {
                                    VSS.getService(VSS.ServiceIds.Navigation).then(function (navigationService) {
                                        navigationService.reload();
                                    });
                                });
                                break;
                            default:
                                alert("Unhandled action: " + command);
                                break;
                        }
                    }
                };
                var menubar = Controls.create(Menus.MenuBar, $(".toolbar"), menubarOptions);
            };
            this.getTabTitle = function (tabContext) {
                _this.log('getTabTitle');
                if (!tabContext || !tabContext.iterationId) {
                    _this.log("getTabTitle: tabContext or tabContext.iterationId empty");
                    return "Goal";
                }
                _this.iterationId = tabContext.iterationId;
                var sprintGoalCookie = _this.getSprintGoalFromCookie();
                if (!sprintGoalCookie) {
                    _this.log("getTabTitle: Sprint goal net yet loaded in cookie, cannot (synchrone) fetch this from storage in 'getTabTitle()' context, call is made anyway");
                    // todo: this call will not return sync. And/but we cannot wait here for the result
                    // because this code run every time the tab is visible (board, capacity, etc.) and we do not want to be blocking and slow down those pages
                    // this way, we at least fetch the values from the server (in the 'background') and persist them in a cookie for the next page view
                    var promise = _this.getSettings(true)
                        .then(function (settings) {
                        // if (settings.sprintGoalInTabLabel && settings.goal != null) {
                        //     return "Goal: " + settings.goal.substr(0, 60);
                        // }
                    });
                    return "Goal";
                }
                if (sprintGoalCookie && sprintGoalCookie.sprintGoalInTabLabel && sprintGoalCookie.goal != null) {
                    _this.log("getTabTitle: loaded title from cookie");
                    return "Goal: " + sprintGoalCookie.goal;
                }
                else {
                    _this.log("getTabTitle: Cookie found but empty goal");
                    return "Goal";
                }
            };
            this.getSprintGoalFromCookie = function () {
                var goal = _this.getCookie(_this.iterationId + _this.teamId + "goalText");
                var sprintGoalInTabLabel = false;
                if (!goal) {
                    goal = _this.getCookie(_this.iterationId + "goalText");
                    sprintGoalInTabLabel = (_this.getCookie(_this.iterationId + "sprintGoalInTabLabel") == "true");
                }
                else {
                    // team specific setting
                    sprintGoalInTabLabel = (_this.getCookie(_this.iterationId + _this.teamId + "sprintGoalInTabLabel") == "true");
                }
                if (!goal)
                    return undefined;
                return {
                    goal: goal,
                    sprintGoalInTabLabel: sprintGoalInTabLabel
                };
            };
            this.saveSettings = function () {
                _this.log('saveSettings');
                if (_this.waitControl)
                    _this.waitControl.startWait();
                var sprintConfig = {
                    sprintGoalInTabLabel: $("#sprintGoalInTabLabel").prop("checked"),
                    goal: $("#goal").val()
                };
                AppInsights.AppInsights.trackEvent("SaveSettings", sprintConfig);
                var configIdentifier = _this.iterationId.toString();
                var configIdentifierWithTeam = _this.iterationId.toString() + _this.teamId;
                _this.updateSprintGoalCookie(configIdentifier, sprintConfig);
                _this.updateSprintGoalCookie(configIdentifierWithTeam, sprintConfig);
                return VSS.getService(VSS.ServiceIds.ExtensionData)
                    .then(function (dataService) {
                    _this.log('saveSettings: ExtensionData Service Loaded');
                    return dataService.setValue("sprintConfig." + configIdentifierWithTeam, sprintConfig).then(function (x) {
                        // override the project level goal, indeed: last team saving 'wins'
                        return dataService.setValue("sprintConfig." + configIdentifier, sprintConfig);
                    });
                })
                    .then(function (value) {
                    _this.log('saveSettings: settings saved!');
                    if (_this.waitControl)
                        _this.waitControl.endWait();
                });
            };
            this.getSettings = function (forceReload) {
                _this.log('getSettings');
                if (_this.waitControl)
                    _this.waitControl.startWait();
                var currentGoalInCookie = _this.getSprintGoalFromCookie();
                var cookieSupport = _this.checkCookie();
                if (forceReload || !currentGoalInCookie || !cookieSupport) {
                    var configIdentifier = _this.iterationId.toString();
                    var configIdentifierWithTeam = _this.iterationId.toString() + _this.teamId;
                    return _this.fetchSettingsFromExtensionDataService(configIdentifierWithTeam).then(function (teamGoal) {
                        if (teamGoal) {
                            _this.updateSprintGoalCookie(configIdentifier, teamGoal);
                            _this.updateSprintGoalCookie(configIdentifierWithTeam, teamGoal);
                            return Q.fcall(function () {
                                // team settings
                                return teamGoal;
                            });
                        }
                        else {
                            // fallback, also for backward compatibility: project/iteration level settings
                            return _this.fetchSettingsFromExtensionDataService(configIdentifier).then(function (iterationGoal) {
                                _this.updateSprintGoalCookie(configIdentifier, iterationGoal);
                                return iterationGoal;
                            });
                        }
                    });
                }
                else {
                    return Q.fcall(function () {
                        _this.log('getSettings: fetched settings from cookie');
                        return currentGoalInCookie;
                    });
                }
            };
            this.fetchSettingsFromExtensionDataService = function (key) {
                return VSS.getService(VSS.ServiceIds.ExtensionData)
                    .then(function (dataService) {
                    _this.log('getSettings: ExtensionData Service Loaded, get value by key: ' + key);
                    return dataService.getValue("sprintConfig." + key);
                })
                    .then(function (sprintGoalDto) {
                    _this.log('getSettings: ExtensionData Service fetched data', sprintGoalDto);
                    if (_this.waitControl)
                        _this.waitControl.endWait();
                    return sprintGoalDto;
                });
            };
            this.updateSprintGoalCookie = function (key, sprintGoal) {
                _this.setCookie(key + "goalText", sprintGoal.goal);
                _this.setCookie(key + "sprintGoalInTabLabel", sprintGoal.sprintGoalInTabLabel);
            };
            this.fillForm = function (sprintGoal) {
                if (!_this.checkCookie()) {
                    $("#cookieWarning").show();
                }
                if (!sprintGoal) {
                    $("#sprintGoalInTabLabel").prop("checked", false);
                    $("#goal").val("");
					$("#goalLabel").text("");
                }
                else {
                    $("#sprintGoalInTabLabel").prop("checked", sprintGoal.sprintGoalInTabLabel);
                    $("#goal").val(sprintGoal.goal);
					$("#goalLabel").text(sprintGoal.goal);
                }
            };
            this.setCookie = function (key, value) {
                var expires = new Date();
                expires.setTime(expires.getTime() + (1 * 24 * 60 * 60 * 1000));
                document.cookie = key + '=' + value + ';expires=' + expires.toUTCString() + ';domain=' + _this.storageUri + ';path=/';
            };
            this.checkCookie = function () {
                _this.setCookie("testcookie", true);
                var success = (_this.getCookie("testcookie") == "true");
                return success;
            };
            this.log = function (message, object) {
                if (object === void 0) { object = null; }
                if (!window.console)
                    return;
                if (_this.storageUri.indexOf('dev') === -1 && _this.storageUri.indexOf('acc') === -1)
                    return;
                if (object) {
                    console.log(message, object);
                    return;
                }
                console.log(message);
            };
            this.loadEmojiPicker = function () {
                _this.addStylesheet('https://maxcdn.bootstrapcdn.com/font-awesome/4.4.0/css/font-awesome.min.css');
                _this.addStylesheet('emojipicker/css/emoji.css');
                _this.addScriptTag('emojipicker/js/config.js');
                _this.addScriptTag('emojipicker/js/util.js');
                _this.addScriptTag('emojipicker/js/jquery.emojiarea.js');
                var emojiPickerScriptElement = _this.addScriptTag('emojipicker/js/emoji-picker.js');
                emojiPickerScriptElement.addEventListener('load', function () {
                    window.emojiPicker = new EmojiPicker({
                        emojiable_selector: '[data-emojiable=true]',
                        assetsPath: 'emojipicker/img',
                        popupButtonClasses: 'fa fa-smile-o'
                    });
                    window.emojiPicker.discover();
                });
            };
            this.addStylesheet = function (href) {
                var link = document.createElement('link');
                link.setAttribute('rel', 'stylesheet');
                link.setAttribute('type', 'text/css');
                link.setAttribute('href', href);
                document.getElementsByTagName('head')[0].appendChild(link);
            };
            this.addScriptTag = function (src) {
                var script = document.createElement('script');
                script.src = src;
                script.async = false;
                document.head.appendChild(script);
                return script;
            };
            var context = VSS.getExtensionContext();
            this.storageUri = this.getLocation(context.baseUri).hostname;
            var webContext = VSS.getWebContext();
            this.teamId = webContext.team.id;
            var config = VSS.getConfiguration();
            this.log('constructor, foregroundInstance = ' + config.foregroundInstance);
            if (config.foregroundInstance) {
                // this code runs when the form is loaded, otherwise, just load the tab
                this.iterationId = config.iterationId;
                this.buildWaitControl();
                this.getSettings(true).then(function (settings) {
                    _this.fillForm(settings);
                    _this.loadEmojiPicker();
                });
                this.buildMenuBar();
                AppInsights.AppInsights.downloadAndSetup({
                    instrumentationKey: "<<AppInsightsInstrumentationKey>>",
                });
                AppInsights.AppInsights.setAuthenticatedUserContext(webContext.user.id, webContext.collection.id);
                AppInsights.AppInsights.trackPageView(document.title, window.location.pathname, {
                    accountName: webContext.account.name,
                    accountId: webContext.account.id,
                    extensionId: context.extensionId,
                    version: context.version
                });
            }
            // register this 'Sprint Goal' service
            VSS.register(VSS.getContribution().id, {
                pageTitle: this.getTabTitle,
                name: this.getTabTitle,
                isInvisible: function (state) {
                    return false;
                }
            });
        }
        SprintGoal.prototype.getCookie = function (key) {
            var keyValue = document.cookie.match('(^|;) ?' + key + '=([^;]*)(;|$)');
            return keyValue ? keyValue[2] : null;
        };
        return SprintGoal;
    }());
    exports.SprintGoal = SprintGoal;
});
//# sourceMappingURL=extension.js.map 
