import * as React from 'react';
import { IEventsProps } from './IEventsProps';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { TextField, PrimaryButton, Checkbox, Stack, Label, Dropdown, IDropdownOption } from '@fluentui/react';
import { DateTimePicker, TimeConvention, TimeDisplayControlType } from '@pnp/spfx-controls-react/lib/DateTimePicker';
import { initializeIcons } from '@fluentui/react/lib/Icons';

initializeIcons();

const CALENDAR_OPTIONS: IDropdownOption[] = [
  { key: 'Albany', text: 'Albany' },
  { key: 'Florida', text: 'Florida' },
  { key: 'Texas', text: 'Texas' }
];

const CATEGORY_OPTIONS: IDropdownOption[] = [
  "Bid Review",
  "Birthday",
  "Event/Conference",
  "Holiday",
  "Home Office (approved)",
  "Interview",
  "Meeting - External",
  "Meeting - Internal",
  "Personal",
  "Service Scheduled",
  "Service Tentative",
  "Site Work",
  "Training"
].map(cat => ({ key: cat, text: cat }));

export interface IEventsState {
  title: string;
  startDateTime: Date;
  endDateTime: Date;
  selectedCalendars: string[];
  selectedCategories: string[];
  location: string;
  description: string;
  addTeamsLink: boolean;
  allDayEvent: boolean;
  loading: boolean;
  hidden: boolean;
}

export default class Events extends React.Component<IEventsProps, IEventsState> {
  constructor(props: IEventsProps) {
    super(props);
    this.state = {
      title: '',
      startDateTime: new Date(),
      endDateTime: new Date(),
      selectedCalendars: [],
      selectedCategories: [],
      location: '',
      description: '',
      addTeamsLink: false,
      allDayEvent: false,
      loading: false,
      hidden: true
    };
  }

  private _onDropdownChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, stateKey?: keyof IEventsState): void => {
    if (option && stateKey) {
      const currentValues = (this.state[stateKey] as string[]) || [];
      let newValues: string[];

      if (option.selected) {
        newValues = [...currentValues, option.key as string];
      } else {
        newValues = currentValues.filter(key => key !== option.key);
      }

      this.setState({ [stateKey as string]: newValues } as any);
    }
  };

  private _triggerFlow = async (): Promise<void> => {
    const { startDateTime, endDateTime, allDayEvent } = this.state;

    if (allDayEvent) {
      const startDateOnly = new Date(startDateTime).setHours(0, 0, 0, 0);
      const endDateOnly = new Date(endDateTime).setHours(0, 0, 0, 0);

      if (endDateOnly < startDateOnly) {
        alert("Error: For All Day events, the End date cannot be before the Start date.");
        return;
      }
    } else {
      if (endDateTime <= startDateTime) {
        alert("Error: The End date/time must be after the Start date/time.");
        return;
      }
    }

    this.setState({ loading: true });

    // Use the exact site URL where the list exists
    const targetSiteUrl = "https://gcontrol.sharepoint.com/sites/test";

    // Encoding the list name for the URL vs metadata
    const listDisplayName = "Event Adder";
    const listInternalName = "Event_x0020_Adder"; // Spaces in internal names become _x0020_

    const nextDay = new Date(this.state.endDateTime);
    nextDay.setDate(nextDay.getDate() + 1);

    const eventData = {
      '__metadata': { 'type': `SP.Data.${listInternalName}ListItem` },
      Title: String(this.state.title),
      Date: this.state.allDayEvent
        ? this.state.startDateTime.toISOString().split('T')[0] + "T00:00:00Z"
        : this.state.startDateTime.toISOString(),
      EndDate: this.state.allDayEvent
        ? nextDay.toISOString().split('T')[0] + "T00:00:00Z"
        : this.state.endDateTime.toISOString(),
      Calendars: String(this.state.selectedCalendars.join(", ")),
      Catergories: String(this.state.selectedCategories.join(", ")),
      EventLocation: String(this.state.location),
      Description: String(this.state.description),
      AddTeamsLink: String(this.state.addTeamsLink),
      AllDayEvent: String(this.state.allDayEvent)
    };

    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-type': 'application/json;odata=verbose',
        'odata-version': ''
      },
      body: JSON.stringify(eventData)
    };

    try {
      // POST to the specific subsite URL
      const response: SPHttpClientResponse = await this.props.context.spHttpClient.post(
        `${targetSiteUrl}/_api/web/lists/getbytitle('${listDisplayName}')/items`,
        SPHttpClient.configurations.v1,
        options
      );

      if (response.ok) {
        alert("Success! Event should be added to calendars shortly...");
      } else {
        const err = await response.text();
        console.error("Detailed Error:", err);
        alert("Still failing...");
      }
    } catch (error) {
      console.error("Request failed:", error);
    } finally {
      this.setState({ loading: false });
    }
  }

  public render(): React.ReactElement<IEventsProps> {
    const stackTokens = { childrenGap: 20 };

    const whiteTextStyles = {
      root: { color: 'white' },
      label: { color: 'white' }, // Targets the labels (Event Title, Location, etc.)
      fieldGroup: { borderColor: 'white' },
      title: { color: 'white', borderColor: 'white', backgroundColor: 'transparent' },
      // Fixes Checkbox hover and checkmark color
      text: { color: 'white' },
      checkbox: { borderColor: 'white' },
      checkmark: { color: 'white' },
      field: {
        color: 'black',
        backgroundColor: 'white'
      },
    };

    const checkboxStyles = {
      ...whiteTextStyles,
      root: {
        selectors: {
          '&:hover .ms-Checkbox-text': { color: 'white' },
          '&:hover .ms-Checkbox-checkbox': { borderColor: 'white' },
        }
      }
    };
    //bruh

    return (
      <section style={{ color: 'white' }}>
        <PrimaryButton
          iconProps={{ iconName: 'Calendar' }}
          text={this.state.hidden ? "Create Event" : "Close"}
          onClick={(e) => this.setState({ hidden: !this.state.hidden })}
          styles={{
            root: { backgroundColor: 'transparent', border: 'none', color: 'white' },
            rootHovered: { backgroundColor: 'transparent', color: 'white' },
            rootPressed: { backgroundColor: 'transparent', color: 'white' },
            icon: { color: 'white' },
            iconHovered: { color: 'white' }
          }}
        />

        {!this.state.hidden && (
          <section style={{ padding: '20px', background: 'transparent', marginTop: '10px' }}>
            <Stack tokens={stackTokens}>
              <Label style={{ fontSize: '20px', borderBottom: '1px solid #ccc', color: 'white' }}>
                Event Details
              </Label>

              <Label style={{ color: 'white' }}>Event Title</Label>
              <TextField
                styles={whiteTextStyles}
                value={this.state.title}
                onChange={(e, val) => this.setState({ title: val || "" })}
              />

              <Label style={{ color: 'white' }}>Enter time in EST, it will update accordingly for Texas</Label>

              <Label style={{ color: 'white' }} >Start Date/Time</Label>
              <DateTimePicker
                showLabels={false}
                value={this.state.startDateTime}
                onChange={(date) => this.setState({ startDateTime: date as Date })}
                timeDisplayControlType={TimeDisplayControlType.Dropdown}
                timeConvention={TimeConvention.Hours12}
              />

              <Label style={{ color: 'white' }}>End Date/Time</Label>
              <DateTimePicker
                showLabels={false}
                value={this.state.endDateTime}
                onChange={(date) => this.setState({ endDateTime: date as Date })}
                timeDisplayControlType={TimeDisplayControlType.Dropdown}
                timeConvention={TimeConvention.Hours12}
              />

              <Checkbox
                label="All day event? (If so dates must match)"
                styles={checkboxStyles}
                checked={this.state.allDayEvent}
                onChange={(e, checked) => this.setState({ allDayEvent: !!checked })}
              />

              <Label style={{ color: 'white' }}>Calendars</Label>
              <Dropdown
                styles={whiteTextStyles}
                multiSelect
                options={CALENDAR_OPTIONS}
                selectedKeys={this.state.selectedCalendars}
                onChange={(e, o) => this._onDropdownChange(e, o, 'selectedCalendars')}
              />

              <Label style={{ color: 'white' }}>Categories</Label>
              <Dropdown
                styles={whiteTextStyles}
                multiSelect
                options={CATEGORY_OPTIONS}
                selectedKeys={this.state.selectedCategories}
                onChange={(e, o) => this._onDropdownChange(e, o, 'selectedCategories')}
              />

              <Label style={{ color: 'white' }}>Location</Label>
              <TextField
                styles={whiteTextStyles}
                value={this.state.location}
                onChange={(e, val) => this.setState({ location: val || "" })}
              />

              <Label style={{ color: 'white' }}>Description</Label>
              <TextField
                styles={whiteTextStyles}
                multiline rows={3}
                value={this.state.description}
                onChange={(e, val) => this.setState({ description: val || "" })}
              />

              <Checkbox
                label="Add Teams link?"
                styles={checkboxStyles}
                checked={this.state.addTeamsLink}
                onChange={(e, checked) => this.setState({ addTeamsLink: !!checked })}
              />

              <PrimaryButton
                text={this.state.loading ? "Creating..." : "Create Event"}
                onClick={this._triggerFlow}
                disabled={this.state.loading || !this.state.title || this.state.selectedCalendars.length === 0}
              />
            </Stack>
          </section>
        )}
      </section>
    );
  }
}
