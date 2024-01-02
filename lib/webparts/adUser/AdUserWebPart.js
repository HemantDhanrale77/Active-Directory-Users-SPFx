var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import { Version } from "@microsoft/sp-core-library";
import { PropertyPaneTextField, } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
// import { escape } from "@microsoft/sp-lodash-subset";
import "datatables.net";
import "datatables.net-dt/css/jquery.dataTables.min.css";
import styles from "./AdUserWebPart.module.scss";
import * as strings from "AdUserWebPartStrings";
import * as $ from "jquery";
import "bootstrap/dist/css/bootstrap.min.css";
import "datatables.net-responsive";
import "datatables.net-responsive-dt/css/responsive.dataTables.min.css";
var AdUserWebPart = /** @class */ (function (_super) {
    __extends(AdUserWebPart, _super);
    function AdUserWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    AdUserWebPart.prototype.render = function () {
        this.domElement.innerHTML = "\n    <section class=\"".concat(styles.adUser, " ").concat(!!this.context.sdks.microsoftTeams ? styles.teams : "", "\">\n      \n      <table id=\"usersTable\" class=\"display table-hover table-striped dt-responsive custom-header\">\n      <thead>\n        <tr>\n          <th>Display Name</th>\n          <th>Description</th>\n          <th>Department</th>\n          <th>Telephone</th>\n          <th>Mobile Phone</th>\n        </tr>\n      </thead>\n      <tbody>\n        <!-- Data will be populated here -->\n      </tbody>\n\n      <tr id=\"loadingRow\">\n            <td colspan=\"8\">\n              <div id=\"loadingIndicator\" class=\"").concat(styles.spinner, "\"></div>\n            </td>\n          </tr>\n      \n    </table>\n    </section>");
        // Initialize DataTable with headers
        $("#usersTable").DataTable({
            responsive: true,
            columns: [
                { title: "Display Name" },
                { title: "Description" },
                { title: "Department" },
                { title: "Telephone" },
                { title: "Mobile Phone" },
                // { title: "User Email" },
            ],
        });
        // Add custom styling for the header
        $(".custom-header th").css("background-color", "#c3c4f3");
        this.getAdUsers();
    };
    AdUserWebPart.prototype.onInit = function () {
        return this._getEnvironmentMessage().then(function (message) {
            // this._environmentMessage = message;
        });
    };
    //Get the Users data from AD
    AdUserWebPart.prototype.getAdUsers = function () {
        return __awaiter(this, void 0, void 0, function () {
            var client, res, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        // Display loading indicator before making the API call
                        $("#loadingIndicator").show();
                        return [4 /*yield*/, this.context.msGraphClientFactory.getClient("3")];
                    case 1:
                        client = _a.sent();
                        return [4 /*yield*/, client
                                .api("/users")
                                .top(999)
                                .select("displayName,jobTitle,mobilePhone,userPrincipalName,department,businessPhones")
                                .get()];
                    case 2:
                        res = _a.sent();
                        // Hide loading indicator after API call
                        $("#loadingIndicator").hide();
                        // Check if valid response received from the API
                        if (res && res.value) {
                            this.populateDataTable(res.value); // Populate the DataTable with the received user data
                        }
                        else {
                            console.log("No user data received from the API");
                        }
                        return [3 /*break*/, 4];
                    case 3:
                        error_1 = _a.sent();
                        // Hide loading indicator in case of an error
                        $("#loadingIndicator").hide();
                        // Log error details to the console
                        console.error("Error fetching users or getting MSGraphClient", error_1);
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    AdUserWebPart.prototype.populateDataTable = function (usersData) {
        // Destroy the existing DataTable (if any)
        var table = $("#usersTable").DataTable();
        table.destroy();
        // Remove the loading row after DataTable initialization
        $("#loadingRow").remove();
        // Initialize DataTable with the new data
        $("#usersTable").DataTable({
            data: usersData,
            responsive: true,
            columns: [
                // Column for displaying user's display name
                {
                    data: "displayName",
                    render: function (data, type, row) {
                        if (data != null && data !== undefined) {
                            return data;
                        }
                        else {
                            return "<div style=\"text-align: center;\">N/A</div>";
                        }
                    },
                },
                // Column for displaying user's job title
                {
                    data: "jobTitle",
                    render: function (data, type, row) {
                        if (data != null && data !== undefined) {
                            return data;
                        }
                        else {
                            return "N/A";
                        }
                    },
                },
                // Column for displaying user's department
                {
                    data: "department",
                    render: function (data, type, row) {
                        if (data != null && data !== undefined) {
                            return data;
                        }
                        else {
                            return "N/A";
                        }
                    },
                },
                // Column for displaying user's business phone (showing the first element of the array)
                {
                    data: "businessPhones",
                    render: function (data, type, row) {
                        if (data && data.length > 0) {
                            return data[0]; // Display the first element of the array
                        }
                        else {
                            return "-"; // Display a hyphen if mobile phone is null or undefined
                        }
                    },
                },
                // { data: "mobilePhone" },
                // Column for displaying user's mobile phone
                {
                    data: "mobilePhone",
                    render: function (data, type, row) {
                        if (data != null && data !== undefined) {
                            return data;
                        }
                        else {
                            return "-"; // Display a hyphen if mobile phone is null or undefined
                        }
                    },
                },
                // { data: "userPrincipalName" },
            ],
        });
    };
    AdUserWebPart.prototype._getEnvironmentMessage = function () {
        var _this = this;
        if (!!this.context.sdks.microsoftTeams) {
            // running in Teams, office.com or Outlook
            return this.context.sdks.microsoftTeams.teamsJs.app
                .getContext()
                .then(function (context) {
                var environmentMessage = "";
                switch (context.app.host.name) {
                    case "Office": // running in Office
                        environmentMessage = _this.context.isServedFromLocalhost
                            ? strings.AppLocalEnvironmentOffice
                            : strings.AppOfficeEnvironment;
                        break;
                    case "Outlook": // running in Outlook
                        environmentMessage = _this.context.isServedFromLocalhost
                            ? strings.AppLocalEnvironmentOutlook
                            : strings.AppOutlookEnvironment;
                        break;
                    case "Teams": // running in Teams
                    case "TeamsModern":
                        environmentMessage = _this.context.isServedFromLocalhost
                            ? strings.AppLocalEnvironmentTeams
                            : strings.AppTeamsTabEnvironment;
                        break;
                    default:
                        environmentMessage = strings.UnknownEnvironment;
                }
                return environmentMessage;
            });
        }
        return Promise.resolve(this.context.isServedFromLocalhost
            ? strings.AppLocalEnvironmentSharePoint
            : strings.AppSharePointEnvironment);
    };
    AdUserWebPart.prototype.onThemeChanged = function (currentTheme) {
        if (!currentTheme) {
            return;
        }
        // this._isDarkTheme = !!currentTheme.isInverted;
        var semanticColors = currentTheme.semanticColors;
        if (semanticColors) {
            this.domElement.style.setProperty("--bodyText", semanticColors.bodyText || null);
            this.domElement.style.setProperty("--link", semanticColors.link || null);
            this.domElement.style.setProperty("--linkHovered", semanticColors.linkHovered || null);
        }
    };
    Object.defineProperty(AdUserWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse("1.0");
        },
        enumerable: false,
        configurable: true
    });
    AdUserWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription,
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField("description", {
                                    label: strings.DescriptionFieldLabel,
                                }),
                            ],
                        },
                    ],
                },
            ],
        };
    };
    return AdUserWebPart;
}(BaseClientSideWebPart));
export default AdUserWebPart;
//# sourceMappingURL=AdUserWebPart.js.map