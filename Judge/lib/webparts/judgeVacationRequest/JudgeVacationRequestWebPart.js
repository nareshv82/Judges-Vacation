var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, PropertyPaneTextField } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient } from '@microsoft/sp-http';
import { sp } from '@pnp/sp';
import { SPComponentLoader } from "@microsoft/sp-loader";
import styles from './JudgeVacationRequestWebPart.module.scss';
import * as strings from 'JudgeVacationRequestWebPartStrings';
var JudgeVacationRequestWebPart = /** @class */ (function (_super) {
    __extends(JudgeVacationRequestWebPart, _super);
    function JudgeVacationRequestWebPart() {
        var _this = _super.call(this) || this;
        _this.listItemEntityTypeName = undefined;
        _this.formatDate = function (date) {
            return new Intl.DateTimeFormat('en-US', {
                year: 'numeric',
                month: 'numeric',
                day: 'numeric'
            })
                .format(new Date(date));
        };
        SPComponentLoader.loadScript('https://code.jquery.com/jquery-1.12.4.min.js', {}).then(function () {
            SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css');
            SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css');
            SPComponentLoader.loadScript('https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js');
            SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.6.4/css/bootstrap-datepicker.min.css');
            SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.6.4/js/bootstrap-datepicker.min.js');
        });
        return _this;
    }
    JudgeVacationRequestWebPart.prototype.onInit = function () {
        return new Promise(function (resolve, reject) {
            sp.setup({
                sp: {
                    headers: {
                        "Accept": "application/json; odata=nometadata"
                    }
                }
            });
            resolve();
        });
    };
    JudgeVacationRequestWebPart.prototype.render = function () {
        this.domElement.innerHTML = "\n      <div class=\"" + styles.judgeVacationRequest + "\">\n        <div class=\"" + styles.container + "\">\n          <div class=\"" + styles.row + "\">\n            <div class=\"" + styles.column + "\">\n              <span class=\"" + styles.title + "\">Welcome to SharePoint!</span>\n              <p class=\"" + styles.subTitle + "\">Customize SharePoint experiences using Web Parts.</p>\n              <p class=\"" + styles.description + "\">" + escape(this.properties.ListName) + "</p>\n              <a href=\"https://aka.ms/spfx\" class=\"" + styles.button + "\">\n                <span class=\"" + styles.label + "\">Learn more</span>\n              </a>\n            </div>\n          </div>\n          <div class=\"row\">\n            <div class=\"col-md-12\">\n                <h1>Judges timekeeping solution</h1>\n                <ul>\n                    <li>\n                        <a href=\"#\" data-toggle=\"modal\" data-target=\"#JudgeVacationModal\" onclick=\"return false\">Judges Vacation Request Form</a>\n                    </li>                  \n                </ul>\n            </div>\n          </div>\n          <div class=\"modal fade in\" id=\"JudgeVacationModal\" role=\"dialog\">\n          <div class=\"modal-dialog modal-xl\" role=\"document\">\n              <div class=\"modal-content\">\n                  <div class=\"form-wrap\">\n                      <div class=\"modal-header\">\n                          <h4 class=\"modal-title thin\">Judge Vacation Request</h4>\n                          <button type=\"button\" class=\"close\" data-dismiss=\"modal\" aria-label=\"Close\" style=\"margin-top:-25px;margin-right:-5px;\">\n                              <span aria-hidden=\"true\">\u00D7</span>\n                          </button>\n                      </div>\n                      <div class=\"modal-body\">\n                          <div class=\"JudgeVacationForm\">\n                              <div class=\"row\">\n                                  <div class=\"col-md-12\">\n                                      <div class=\"row\">\n                                          <div class=\"col-md-12\">\n                                              <div class=\"form-group\">\n                                                  <label class=\"control-label\">Requester:</label>\n                                                  <div class=\"input-group\">\n                                                      <span class=\"input-group-addon\">\n                                                          <i class=\"fa fa-user\" aria-hidden=\"true\"></i>\n                                                      </span>\n                                                      <input type=\"text\" class=\"form-control\" placeholder=\"Enter Requester Name\" data-info=\"RequesterName\" data-error=\"Enter a Requester Name!\" required=\"\" id=\"requesterName\">\n                                                  </div>\n                                                  <span class=\"glyphicon form-control-feedback\" aria-hidden=\"true\"></span>\n                                                  <div class=\"help-block with-errors\"></div>\n                                              </div>\n                                          </div>                                    \n                                      </div>\n\n                                      <div class=\"row\">\n                                      <div class=\"col-sm-6\">\n                                      <div class=\"input-group date\" data-provide=\"datepicker\">\n                                          <input type=\"text\" class=\"form-control\"  placeholder=\"Enter a Start Date\"  id=\"startDate\">\n                                          <div class=\"input-group-addon\">\n                                              <span class=\"glyphicon glyphicon-calendar\"></span>\n                                          </div>\n                                      </div>\n                                  </div>\n                                          <div class=\"col-md-12\">\n                                              <div class=\"form-group\">\n                                                  <label class=\"control-label\">Request Start Date:</label>\n                                                  <div class=\"input-group date\" data-provide=\"datepicker\">\n                                                      <input type=\"text\" class=\"form-control\"   id=\"startDate\">\n                                                      <div class=\"input-group-addon\">\n                                                        <span class=\"glyphicon glyphicon-calendar\"></span>\n                                                      </div>\n                                                  </div>\n                                                  <span class=\"glyphicon form-control-feedback\" style=\"margin-right:40px;\" aria-hidden=\"true\"></span>\n                                                  <div class=\"help-block with-errors\"></div>\n                                              </div>\n                                          </div>\n                                      </div>\n                                        <div class=\"row\">\n                                          <div class=\"col-md-12\">\n                                              <div class=\"form-group\">\n                                                  <label class=\"control-label\">Request End Date:</label>\n                                                  <div class=\"input-group date\" data-provide=\"datepicker\">\n                                                      <input type=\"text\" class=\"form-control\" placeholder=\"Enter a End Date\" id=\"endDate\">\n                                                      <span class=\"input-group-addon\"><span class=\"glyphicon glyphicon-calendar\"></span></span>\n                                                  </div>\n                                                  <span class=\"glyphicon form-control-feedback\" style=\"margin-right:40px;\" aria-hidden=\"true\"></span>\n                                                  <div class=\"help-block with-errors\"></div>\n                                              </div>\n                                          </div>\n                                      </div>\n                                      <div class=\"row\">\n                                          <div class=\"col-md-12\">                                        \n                                              <div class=\"form-group has-feedback\">\n                                                  <label class=\"control-label\">Amount of Time:</label>\n                                                  <select class=\"form-control select2-hidden-accessible\" data-info=\"AmountofTime\" required=\"\" tabindex=\"-1\" aria-hidden=\"true\" id=\"dayOff\">\n                                                      <option value=\"Full day\">Full day</option>\n                                                      <option value=\"Half day am\">Half day am</option>\n                                                      <option value=\"Half day pm\">Half day pm</option>\n                                                  </select>\t\t\t\t\t\t\t\t\t\t\t\n                                                  <span class=\"glyphicon form-control-feedback shiftLeft\" aria-hidden=\"true\"></span>\n                                                  <div class=\"help-block with-errors\"></div>\n                                              </div>\n                                          </div>                                    \n                                      </div>    \n                                      <div class=\"row\">\n                                          <div class=\"col-md-12\">                                        \n                                              <div class=\"form-group has-feedback\">\n                                                  <label class=\"control-label\">Request Type:</label>\n                                                  <select class=\"form-control select2-hidden-accessible\" data-info=\"RequestType\" required=\"\" tabindex=\"-1\" aria-hidden=\"true\" id=\"requestType\">\n                                                      <option value=\"Vacation\">Vacation</option>\n                                                      <option value=\"Sick\">Sick</option>\n                                                  </select>\t\t\t\t\t\t\t\t\t\t\t\n                                                  <span class=\"glyphicon form-control-feedback shiftLeft\" aria-hidden=\"true\"></span>\n                                                  <div class=\"help-block with-errors\"></div>\n                                              </div>\n                                          </div>                                    \n                                      </div>\t\t\t\t\t\t\t\t\n                                      <div class=\"row\">\n                                          <div class=\"col-md-12\">\n                                              <div class=\"form-group has-feedback\">\n                                                  <label class=\"control-label\">Comments:</label>\n                                                  <div class=\"input-group\">\n                                                      <span class=\"input-group-addon\">\n                                                          <i class=\"fa fa-pencil\"></i>\n                                                      </span>\n                                                      <textarea class=\"form-control\" rows=\"3\" placeholder=\"Enter Comments...\" data-info=\"Comments\" id=\"requesterComments\"></textarea>\n                                                  </div>\n                                                  <span class=\"glyphicon form-control-feedback\" aria-hidden=\"true\"></span>\n                                                  <div class=\"help-block with-errors\"></div>\n                                              </div>\n                                          </div>\n                                      </div>\n\n                                  </div>\n                              </div>\n                          </div>\n                      </div>\n                      <div class=\"modal-footer\">\n                          <div class=\"form-group\">\n                              <button type=\"submit\" class=\"btn btn-primary create-Button\"> <i class=\"fa fa-check\" aria-hidden=\"true\"></i>\n                                  Submit</button>\n                              <button class=\"btn btn-default\" data-dismiss=\"modal\"> <i class=\"fa fa-times\" aria-hidden=\"true\"></i> Cancel</button>\n                          </div>\n                          <div class=\"status\"></div>\n                          <div class=\"items\"></div>\n                      </div>\n                      \n                  </div>\n              </div>\n          </div>\n      </div>\n      <script type=\"text/javascritp\">\n      $(document).ready(function(){\n        $('#startDate').datepicker();\n        $('#endDate').datepicker();\n      });\n      </script>\n        </div>\n      </div>";
        this.updateStatus(this.listNotConfigured() ? 'Please configure list in Web Part properties' : 'Ready');
        this.setButtonsState();
        this.setButtonsEventHandlers();
        this.getSPData();
    };
    JudgeVacationRequestWebPart.prototype.setButtonsState = function () {
        var buttons = this.domElement.querySelectorAll("button." + styles.button);
        var listNotConfigured = this.listNotConfigured();
        for (var i = 0; i < buttons.length; i++) {
            var button = buttons.item(i);
            if (listNotConfigured) {
                button.setAttribute('disabled', 'disabled');
            }
            else {
                button.removeAttribute('disabled');
            }
        }
    };
    JudgeVacationRequestWebPart.prototype.getSPData = function () {
        var _this = this;
        sp.profiles.myProperties.get().then(function (r) {
            console.log(r);
            var dept = r["UserProfileProperties"][11]["Value"];
            var mgr = r["UserProfileProperties"][14]["Value"];
            var mgrID;
            var payload = JSON.stringify({
                'logonName': mgr //this.context.pageContext.user.loginName // i:0#.f|membership|firstname.lastname@contoso.onmicrosoft.com      
            });
            var postData = {
                body: payload
            };
            var endPoint = _this.context.pageContext.site.absoluteUrl + "/_api/web/ensureuser";
            _this.context.spHttpClient.post(endPoint, SPHttpClient.configurations.v1, postData)
                .then(function (response) {
                response.json().then(function (resposneJSON) {
                    console.log("manager");
                    console.log(resposneJSON);
                    _this.Manager = resposneJSON.Id;
                });
                //return response.json();
            });
            _this.renderData(r['DisplayName'], dept, mgrID);
        });
    };
    JudgeVacationRequestWebPart.prototype.renderData = function (strResponse, strDept, strMgr) {
        document.getElementById("requesterName")["value"] = strResponse;
        document.getElementById("startDate")["value"] = this.formatDate(new Date().toString());
        document.getElementById("endDate")["value"] = this.formatDate(new Date().toString());
        this.Department = strDept;
        //this.Manager=strMgr;
    };
    JudgeVacationRequestWebPart.prototype.setButtonsEventHandlers = function () {
        var webPart = this;
        this.domElement.querySelector('button.create-Button').addEventListener('click', function () { webPart.createItem(); });
        //this.domElement.querySelector('button.read-Button').addEventListener('click', () => { webPart.readItem(); });
        //this.domElement.querySelector('button.readall-Button').addEventListener('click', () => { webPart.readItems(); });
        //this.domElement.querySelector('button.update-Button').addEventListener('click', () => { webPart.updateItem(); });
        //this.domElement.querySelector('button.delete-Button').addEventListener('click', () => { webPart.deleteItem(); });
    };
    Object.defineProperty(JudgeVacationRequestWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    JudgeVacationRequestWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('description', {
                                    label: strings.ListNameFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    JudgeVacationRequestWebPart.prototype.listNotConfigured = function () {
        return this.properties.ListName === undefined ||
            this.properties.ListName === null ||
            this.properties.ListName.length === 0;
    };
    JudgeVacationRequestWebPart.prototype.createItem = function () {
        var _this = this;
        this.updateStatus('Creating item...');
        sp.web.currentUser.get().then(function (res) {
            console.log(res.Id);
            sp.web.lists.getByTitle("Judges Vacation Calendar").items.add({
                'Title': "Out of Office - " + document.getElementById("requesterName")["value"],
                'RequesterId': res.Id,
                'StartDate': new Date(document.getElementById("startDate")["value"]),
                'EndDate': new Date(document.getElementById("endDate")["value"] + " 11:45:00 PM"),
                //'HalfDay': document.getElementById("dayOff")["value"],
                //'Vacation_x0020_Type':document.getElementById("requestType")["value"],
                //'Description':document.getElementById("requesterComments"),
                'Department': _this.Department,
                'ManagerId': _this.Manager
            }).then(function (result) {
                var item = result.data;
                _this.updateStatus("Item '" + item.Title + "' (ID: " + item.Id + ") successfully created");
            }, function (error) {
                _this.updateStatus('Error while creating the item: ' + error);
            });
        });
    };
    JudgeVacationRequestWebPart.prototype.readItem = function () {
        var _this = this;
        this.updateStatus('Loading latest items...');
        this.getLatestItemId()
            .then(function (itemId) {
            if (itemId === -1) {
                throw new Error('No items found in the list');
            }
            _this.updateStatus("Loading information about item ID: " + itemId + "...");
            return sp.web.lists.getByTitle(_this.properties.ListName)
                .items.getById(itemId).select('Title', 'Id').get();
        })
            .then(function (item) {
            _this.updateStatus("Item ID: " + item.Id + ", Title: " + item.Title);
        }, function (error) {
            _this.updateStatus('Loading latest item failed with error: ' + error);
        });
    };
    JudgeVacationRequestWebPart.prototype.getLatestItemId = function () {
        var _this = this;
        return new Promise(function (resolve, reject) {
            sp.web.lists.getByTitle(_this.properties.ListName)
                .items.orderBy('Id', false).top(1).select('Id').get()
                .then(function (items) {
                if (items.length === 0) {
                    resolve(-1);
                }
                else {
                    resolve(items[0].Id);
                }
            }, function (error) {
                reject(error);
            });
        });
    };
    JudgeVacationRequestWebPart.prototype.readItems = function () {
        var _this = this;
        this.updateStatus('Loading all items...');
        sp.web.lists.getByTitle(this.properties.ListName)
            .items.select('Title', 'Id').get()
            .then(function (items) {
            _this.updateStatus("Successfully loaded " + items.length + " items", items);
        }, function (error) {
            _this.updateStatus('Loading all items failed with error: ' + error);
        });
    };
    JudgeVacationRequestWebPart.prototype.updateItem = function () {
        var _this = this;
        this.updateStatus('Loading latest items...');
        var latestItemId = undefined;
        var etag = undefined;
        this.getLatestItemId()
            .then(function (itemId) {
            if (itemId === -1) {
                throw new Error('No items found in the list');
            }
            latestItemId = itemId;
            _this.updateStatus("Loading information about item ID: " + itemId + "...");
            return sp.web.lists.getByTitle(_this.properties.ListName)
                .items.getById(itemId).get(undefined, {
                headers: {
                    'Accept': 'application/json;odata=minimalmetadata'
                }
            });
        })
            .then(function (item) {
            etag = item["odata.etag"];
            return Promise.resolve(item);
        })
            .then(function (item) {
            return sp.web.lists.getByTitle(_this.properties.ListName)
                .items.getById(item.Id).update({
                'Title': "Item " + new Date()
            }, etag);
        })
            .then(function (result) {
            _this.updateStatus("Item with ID: " + latestItemId + " successfully updated");
        }, function (error) {
            _this.updateStatus('Loading latest item failed with error: ' + error);
        });
    };
    JudgeVacationRequestWebPart.prototype.deleteItem = function () {
        var _this = this;
        if (!window.confirm('Are you sure you want to delete the latest item?')) {
            return;
        }
        this.updateStatus('Loading latest items...');
        var latestItemId = undefined;
        var etag = undefined;
        this.getLatestItemId()
            .then(function (itemId) {
            if (itemId === -1) {
                throw new Error('No items found in the list');
            }
            latestItemId = itemId;
            _this.updateStatus("Loading information about item ID: " + latestItemId + "...");
            return sp.web.lists.getByTitle(_this.properties.ListName)
                .items.getById(latestItemId).select('Id').get(undefined, {
                headers: {
                    'Accept': 'application/json;odata=minimalmetadata'
                }
            });
        })
            .then(function (item) {
            etag = item["odata.etag"];
            return Promise.resolve(item);
        })
            .then(function (item) {
            _this.updateStatus("Deleting item with ID: " + latestItemId + "...");
            return sp.web.lists.getByTitle(_this.properties.ListName)
                .items.getById(item.Id).delete(etag);
        })
            .then(function () {
            _this.updateStatus("Item with ID: " + latestItemId + " successfully deleted");
        }, function (error) {
            _this.updateStatus("Error deleting item: " + error);
        });
    };
    JudgeVacationRequestWebPart.prototype.updateStatus = function (status, items) {
        if (items === void 0) { items = []; }
        this.domElement.querySelector('.status').innerHTML = status;
        this.updateItemsHtml(items);
    };
    JudgeVacationRequestWebPart.prototype.updateItemsHtml = function (items) {
        this.domElement.querySelector('.items').innerHTML = items.map(function (item) { return "<li>" + item.Title + " (" + item.Id + ")</li>"; }).join("");
    };
    return JudgeVacationRequestWebPart;
}(BaseClientSideWebPart));
export default JudgeVacationRequestWebPart;
//# sourceMappingURL=JudgeVacationRequestWebPart.js.map