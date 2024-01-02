import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import type { IReadonlyTheme } from "@microsoft/sp-component-base";
// import { escape } from "@microsoft/sp-lodash-subset";
import "datatables.net";
import "datatables.net-dt/css/jquery.dataTables.min.css";

import styles from "./AdUserWebPart.module.scss";
import * as strings from "AdUserWebPartStrings";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import * as $ from "jquery";
import "bootstrap/dist/css/bootstrap.min.css";

import "datatables.net-responsive";
import "datatables.net-responsive-dt/css/responsive.dataTables.min.css";

export interface IAdUserWebPartProps {
  description: string;
}

export default class AdUserWebPart extends BaseClientSideWebPart<IAdUserWebPartProps> {
  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.adUser} ${
      !!this.context.sdks.microsoftTeams ? styles.teams : ""
    }">
      
      <table id="usersTable" class="display table-hover table-striped dt-responsive custom-header">
      <thead>
        <tr>
          <th>Display Name</th>
          <th>Description</th>
          <th>Department</th>
          <th>Telephone</th>
          <th>Mobile Phone</th>
        </tr>
      </thead>
      <tbody>
        <!-- Data will be populated here -->
      </tbody>

      <tr id="loadingRow">
            <td colspan="8">
              <div id="loadingIndicator" class="${styles.spinner}"></div>
            </td>
          </tr>
      
    </table>
    </section>`;

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
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then((message) => {
      // this._environmentMessage = message;
    });
  }

  //Get the Users data from AD
  private async getAdUsers() {
    try {
      // Display loading indicator before making the API call
      $("#loadingIndicator").show();

      // Get the MS Graph client using the context
      const client: MSGraphClientV3 =
        await this.context.msGraphClientFactory.getClient("3");

      // Make the API call to retrieve user data
      const res = await client
        .api("/users")
        .top(999)
        .select(
          "displayName,jobTitle,mobilePhone,userPrincipalName,department,businessPhones"
        )
        .get();

      // Hide loading indicator after API call
      $("#loadingIndicator").hide();

      // Check if valid response received from the API
      if (res && res.value) {
        this.populateDataTable(res.value); // Populate the DataTable with the received user data
      } else {
        console.log("No user data received from the API");
      }
    } catch (error) {
      // Hide loading indicator in case of an error
      $("#loadingIndicator").hide();

      // Log error details to the console
      console.error("Error fetching users or getting MSGraphClient", error);
    }
  }

  private populateDataTable(usersData: any) {
    // Destroy the existing DataTable (if any)
    const table = $("#usersTable").DataTable();
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
            } else {
              return `<div style="text-align: center;">N/A</div>`;
            }
          },
        },

        // Column for displaying user's job title
        {
          data: "jobTitle",
          render: function (data, type, row) {
            if (data != null && data !== undefined) {
              return data;
            } else {
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
            } else {
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
            } else {
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
            } else {
              return "-"; // Display a hyphen if mobile phone is null or undefined
            }
          },
        },
        // { data: "userPrincipalName" },
      ],
    });
  }
  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
            case "TeamsModern":
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    // this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
  }
}
