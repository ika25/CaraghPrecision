import * as React from 'react';
import styles from './EmployeeDirectory.module.scss';
import type { IEmployeeDirectoryProps } from './IEmployeeDirectoryProps';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

interface IPhoneDirectoryListItem {
  Id: number;
  Title: string;
  Department?: string;
  Extension?: string;
  Mobile?: string;
  ContactType?: string;
  Active?: boolean;
  JobTitle?: string;
  Email?: string;
}

interface IEmployee {
  id: number;
  name: string;
  jobTitle: string;
  department: string;
  extension: string;
  phone: string;
  email: string;
  contactType: string;
}

interface IEmployeeDirectoryState {
  employees: IEmployee[];
  searchText: string;
  department: string;
  contactType: string;
  isLoading: boolean;
  errorMessage: string;
}

export default class EmployeeDirectory extends React.Component<
  IEmployeeDirectoryProps,
  IEmployeeDirectoryState
> {
  public constructor(props: IEmployeeDirectoryProps) {
    super(props);

    this.state = {
      employees: [],
      searchText: '',
      department: 'All',
      contactType: 'All',
      isLoading: true,
      errorMessage: ''
    };
  }

  public componentDidMount(): void {
    void this._loadEmployees();
  }

  private async _loadEmployees(): Promise<void> {
    try {
      const webUrl = this.props.context.pageContext.web.absoluteUrl;

      const endpoint =
        `${webUrl}/_api/web/lists/getbytitle('Phone Directory')/items` +
        `?$select=Id,Title,field_1,field_2,field_3,field_4,field_5,JobTitle,Email` +
        `&$filter=field_5 eq 'Yes'` +
        `&$orderby=Title asc` +
        `&$top=5000`;

      const response: SPHttpClientResponse =
        await this.props.context.spHttpClient.get(
          endpoint,
          SPHttpClient.configurations.v1
        );

      if (!response.ok) {
        throw new Error(
          `SharePoint returned ${response.status}: ${response.statusText}`
        );
      }

      const data: {
        value: Array<{
          Id: number;
          Title: string;
          field_1?: string;
          field_2?: number;
          field_3?: string;
          field_4?: string;
          field_5?: string;
          JobTitle?: string;
          Email?: string;
        }>;
      } = await response.json();

      const employees: IEmployee[] = data.value.map(item => ({
        id: item.Id,
        name: item.Title || 'Unnamed contact',
        jobTitle: item.JobTitle || '',
        department: item.field_1 || '',
        extension: item.field_2 ? String(item.field_2) : '',
        phone: item.field_3 || '',
        email: item.Email || '',
        contactType: item.field_4 || ''
      }));

      this.setState({
        employees,
        isLoading: false,
        errorMessage: ''
      });
    } catch (error) {
      const errorMessage =
        error instanceof Error
          ? error.message
          : 'An unexpected error occurred while loading the directory.';

      this.setState({
        employees: [],
        isLoading: false,
        errorMessage
      });
    }
  }

  public render(): React.ReactElement<IEmployeeDirectoryProps> {
    const departments: string[] = [
      'All',
      ...Array.from(
        new Set(
          this.state.employees
            .map(employee => employee.department)
            .filter(department => department.length > 0)
        )
      ).sort()
    ];

    const contactTypes: string[] = [
      'All',
      ...Array.from(
        new Set(
          this.state.employees
            .map(employee => employee.contactType)
            .filter(contactType => contactType.length > 0)
        )
      ).sort()
    ];

    const search = this.state.searchText.trim().toLowerCase();

    const filteredEmployees = this.state.employees.filter(employee => {
      const matchesSearch =
        search.length === 0 ||
        employee.name.toLowerCase().includes(search) ||
        employee.jobTitle.toLowerCase().includes(search) ||
        employee.department.toLowerCase().includes(search) ||
        employee.extension.toLowerCase().includes(search) ||
        employee.phone.toLowerCase().includes(search) ||
        employee.email.toLowerCase().includes(search) ||
        employee.contactType.toLowerCase().includes(search);

      const matchesDepartment =
        this.state.department === 'All' ||
        employee.department === this.state.department;

      const matchesContactType =
        this.state.contactType === 'All' ||
        employee.contactType === this.state.contactType;

      return matchesSearch && matchesDepartment && matchesContactType;
    });

    return (
      <section className={styles.employeeDirectory}>
        <div className={styles.header}>
          <h2>Employee &amp; Phone Directory</h2>
          <p>
            Find Caragh employees, rooms, mobile numbers and extensions.
          </p>
        </div>

        <div className={styles.filters}>
          <input
            type="search"
            aria-label="Search the phone directory"
            placeholder="Search name, department or extension..."
            value={this.state.searchText}
            onChange={(event: React.ChangeEvent<HTMLInputElement>) =>
              this.setState({ searchText: event.target.value })
            }
          />

          <select
            aria-label="Filter by department"
            value={this.state.department}
            onChange={(event: React.ChangeEvent<HTMLSelectElement>) =>
              this.setState({ department: event.target.value })
            }
          >
            {departments.map(department => (
              <option key={department} value={department}>
                {department === 'All'
                  ? 'All departments'
                  : department}
              </option>
            ))}
          </select>

          <select
            aria-label="Filter by contact type"
            value={this.state.contactType}
            onChange={(event: React.ChangeEvent<HTMLSelectElement>) =>
              this.setState({ contactType: event.target.value })
            }
          >
            {contactTypes.map(contactType => (
              <option key={contactType} value={contactType}>
                {contactType === 'All'
                  ? 'All contact types'
                  : contactType}
              </option>
            ))}
          </select>
        </div>

        {this.state.isLoading && (
          <div className={styles.noResults}>
            Loading Phone Directory...
          </div>
        )}
        <div className={styles.resultsBar}>
          <span>
            {filteredEmployees.length}{' '}
            {filteredEmployees.length === 1 ? 'contact' : 'contacts'}
          </span>

          {(this.state.searchText ||
            this.state.department !== 'All' ||
            this.state.contactType !== 'All') && (
              <button
                type="button"
                onClick={() =>
                  this.setState({
                    searchText: '',
                    department: 'All',
                    contactType: 'All'
                  })
                }
              >
                Clear filters
              </button>
            )}
        </div>
        {!this.state.isLoading &&
          this.state.errorMessage.length > 0 && (
            <div className={styles.noResults}>
              <strong>Unable to load the Phone Directory.</strong>
              <br />
              {this.state.errorMessage}
            </div>
          )}

        {!this.state.isLoading &&
          this.state.errorMessage.length === 0 && (
            <div className={styles.results}>
              {filteredEmployees.map(employee => (
                <div
                  className={styles.employeeCard}
                  key={employee.id}
                >
                  <div className={styles.avatar}>
                    {employee.name
                      .split(' ')
                      .filter(part => part.length > 0)
                      .map(part => part.charAt(0))
                      .join('')
                      .substring(0, 2)
                      .toUpperCase()}
                  </div>

                  <div className={styles.employeeInfo}>
                    <h3>{employee.name}</h3>

                    {employee.jobTitle && (
                      <div className={styles.jobTitle}>
                        {employee.jobTitle}
                      </div>
                    )}

                    <div className={styles.details}>
                      {employee.department && (
                        <span>{employee.department}</span>
                      )}

                      {employee.contactType && (
                        <span>{employee.contactType}</span>
                      )}
                    </div>

                    <div className={styles.contact}>
                      {employee.phone && (
                        <a
                          href={`tel:${employee.phone}`}
                          title={`Call ${employee.name}`}
                        >
                          ☎ {employee.phone}
                        </a>
                      )}

                      {employee.extension && (
                        <span className={styles.extension}>
                          Ext. {employee.extension}
                        </span>
                      )}

                      {employee.email && (
                        <a
                          href={`mailto:${employee.email}`}
                          title={`Email ${employee.name}`}
                        >
                          ✉ {employee.email}
                        </a>
                      )}

                    </div>
                  </div>
                </div>
              ))}

              {filteredEmployees.length === 0 && (
                <div className={styles.noResults}>
                  No matching directory contacts found.
                </div>
              )}
            </div>
          )}
      </section>
    );
  }
}