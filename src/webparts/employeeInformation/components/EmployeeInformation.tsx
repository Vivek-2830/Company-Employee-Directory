import * as React from 'react';
import styles from './EmployeeInformation.module.scss';
import type { IEmployeeInformationProps } from './IEmployeeInformationProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/attachments";
import { IFieldInfo } from "@pnp/sp/fields";
import { Dropdown, SearchBox } from '@fluentui/react';
import { IItem, Item } from '@pnp/sp/items';
import { Attachment, Attachments, IAttachmentInfo } from '@pnp/sp/attachments';

export interface IEmployeeInformationState {
  Teamlist: any;
  Team: any;
  EmployeeDetails: any;
  AllEmployeeDetails: any;
}

export default class EmployeeInformation extends React.Component<IEmployeeInformationProps, IEmployeeInformationState> {

  private sp: any;

  constructor(props: IEmployeeInformationProps, state: IEmployeeInformationState) {

    super(props);

    this.state = {
      Teamlist: [],
      Team: [],
      EmployeeDetails: "",
      AllEmployeeDetails: ""
    };

    this.sp = spfi().using(SPFx(this.props.context));

  }
  public render(): React.ReactElement<IEmployeeInformationProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      userDisplayName
    } = this.props;

    return (
      <section className="employeeInformation">

        <section className={styles.companyDirectory}>

          <div className={styles.directoryHeader}>

            <div className={styles.headerContent}>
              <h1>Company Directory</h1>

              <p>
                Connect with our global talent across every Team. Discover expertise,
                <br />
                foster collaboration, and build your internal network.
              </p>
            </div>

          </div>

          <div className={styles.filterContainer}>

            <div>

              <SearchBox className={styles['new-search']}
                type="text"
                placeholder="Search by name, expertise, or keywords..."
                onChange={(e: any) => { this.applyVendorFilters(e.target.value); }}
                onClear={(e: any) => { this.applyVendorFilters(e.target.value); }}
              />
            </div>

            <div>
              <Dropdown
                className={styles.newDropdown}
                options={this.state.Teamlist}
                label="Team"
                required
                placeholder="Select Team"
                selectedKey={this.state.Team}
                onChange={(event, option) => {
                  if (option) {
                    this.setState({
                      Team: option.text
                    });
                  }
                }}
              />
            </div>

          </div>

          {
            this.state.EmployeeDetails.length > 0 &&
            this.state.EmployeeDetails.map((item: any) => {
              return (
                <div
                  className={styles.employeeGrid}
                  key={item.ID}
                >
                  <div className={styles.employeeCard}>

                    {/* Employee Image */}
                    <div className={styles.employeeImageWrapper}>
                      {item.ImageUrl ? (
                        <img
                          src={item.ImageUrl}
                          className={styles.employeeImage}
                          alt={item.Title || "Employee"}
                        />
                      ) : (
                        <div className={styles.employeeInfo}>
                          {item.Title
                            ? item.Title.charAt(0).toUpperCase()
                            : "E"}
                        </div>
                      )}
                    </div>

                    {/* Employee Information */}
                    <div className={styles.employeeInfo}>

                      <h3>
                        {item.Title}
                      </h3>

                      <div className={styles.employeeRole}>
                        {item.Designation}
                      </div>

                      <div className={styles.employeeLocation}>
                        {item.Department}
                        {item.Department && item.Team ? " • " : ""}
                        {item.Team}
                      </div>

                    </div>

                    {/* Employee Contact */}
                    <div className={styles.employeeContact}>

                      <div className={styles.contactItem}>
                        <span className={styles.contactIcon}>
                          ✉
                        </span>

                        <span>
                          {item.EmailId}
                        </span>
                      </div>

                      <div className={styles.contactItem}>
                        <span className={styles.contactIcon}>
                          ♧
                        </span>

                        <span>
                          {item.DOB}
                        </span>
                      </div>

                    </div>

                  </div>
                </div>
              );
            })
          }

        </section>

      </section>
    );
  }

  public async componentDidMount() {
    this.GetEmployeeDetails();
    this.getEmployeeData();
  }

  public async getEmployeeData(): Promise<void> {
    try {
      const items = await this.sp.web.lists
        .getByTitle("Employee Information")
        .items
        .select(
          "ID",
          "Title",
          "Designation",
          "Department",
          "Team",
          "MobileNumber",
          "EmailId",
          "DOB",
          "DOJ",
          "LWD",
          "ReportsTo",
          "CurrentAddress",
          "PermanentAddress",
          "Gender",
          "MaritalStatus",
          "BloodGroup",
          "PersonalEmail",
          "EmployeeNumber",
          "ProfilePictureName",
          "SkypeID",
          "Status",
          "TeamLead"
        ).expand("ReportsTo" , "TeamLead")();

      console.log("Employee Data:", items);

      const AllData = await Promise.all(
        items.map(async (item: any) => {

          let info: IAttachmentInfo[] = [];

          // Get attachments for current employee
          try {
            info = await this.sp.web.lists
              .getByTitle("Employee Information")
              .items
              .getById(item.ID)
              .attachmentFiles();

            console.log(
              `Attachments for Employee ID ${item.ID}:`,
              info
            );

          } catch (attachmentError) {
            console.error(
              `Error fetching attachments for employee ID ${item.ID}:`,
              attachmentError
            );
          }

          // Get first attachment as employee image
          const ImageUrl =
            info.length > 0
              ? info[0].ServerRelativeUrl
              : "";

          return {
            ID: item.ID || "",
            Title: item.Title || "",
            Designation: item.Designation || "",
            Department: item.Department || "",
            Team: item.Team || "",
            MobileNumber: item.MobileNumber || "",
            EmailId: item.EmailId || "",
            DOB: item.DOB || "",
            DOJ: item.DOJ || "",
            LWD: item.LWD || "",

            ReportsTo: item.ReportsTo || "",

            CurrentAddress: item.CurrentAddress || "",
            PermanentAddress: item.PermanentAddress || "",
            Gender: item.Gender || "",
            MaritalStatus: item.MaritalStatus || "",
            BloodGroup: item.BloodGroup || "",
            PersonalEmail: item.PersonalEmail || "",
            EmployeeNumber: item.EmployeeNumber || "",
            ProfilePictureName: item.ProfilePictureName || "",
            SkypeID: item.SkypeID || "",
            Status: item.Status || "",

            TeamLead: item.TeamLead || "",

            // All attachments
            Attachments: info,

            // First attachment used as profile image
            ImageUrl: ImageUrl
          };
        })
      );

      this.setState({
        EmployeeDetails: AllData
      });

      console.log("All Employee Details:", AllData);

    } catch (error) {
      console.error("Error fetching employee data:", error);
    }
  }

  public async GetEmployeeDetails(): Promise<void> {
    try {
      const choiceFieldName = "Team";

      const field1: IFieldInfo = await this.sp.web.lists
        .getByTitle("Employee Information")
        .fields
        .getByInternalNameOrTitle(choiceFieldName)();

      const Teamlist: { key: number; text: string }[] = [];

      if (field1.Choices) {
        field1.Choices.forEach((dname: string, i: number) => {
          Teamlist.push({
            key: i,
            text: dname
          });
        });
      }

      this.setState({
        Teamlist: Teamlist
      });

      console.log("Field Details:", field1);
      console.log("Team List:", Teamlist);

    } catch (error) {
      console.error("Error fetching Team field:", error);
    }
  }

  private async applyVendorFilters(Test: string): Promise<void> {
    if (Test) {
      let SerchText = Test.toLowerCase();

      let filteredData = this.state.AllEmployeeDetails.filter((x: any) => {
        let ProjectName = x.ProjectName.toLowerCase();
        let ProjectManager = x.ProjectManager.toLowerCase();
        return (
          ProjectName.includes(SerchText) || ProjectManager.includes(SerchText)
        );
      });

      this.setState({ EmployeeDetails: filteredData });
    }
    else {
      this.setState({ EmployeeDetails: this.state.AllEmployeeDetails });
    }
  }

}
