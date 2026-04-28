import * as React from 'react';
import { IEventsProps } from './IEventsProps';
import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
import { TextField, PrimaryButton, Checkbox, Stack, Label, Dropdown, IDropdownOption } from '@fluentui/react';
import { DateTimePicker } from '@pnp/spfx-controls-react/lib/DateTimePicker';

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
  loading: boolean;
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
      loading: false
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
    this.setState({ loading: true });

    // PASTE YOUR FLOW URL HERE
    const flowUrl = "YOUR_POWER_AUTOMATE_HTTP_URL_HERE";

    const eventData = {
      title: this.state.title,
      startDateTime: this.state.startDateTime.toISOString(),
      endDateTime: this.state.endDateTime.toISOString(),
      calendars: this.state.selectedCalendars,
      categories: this.state.selectedCategories,
      location: this.state.location,
      description: this.state.description,
      addTeamsLink: this.state.addTeamsLink,
    };

    const options: IHttpClientOptions = {
      body: JSON.stringify(eventData),
      headers: { 'Content-Type': 'application/json' }
    };

    try {
      const response: HttpClientResponse = await this.props.context.httpClient.post(
        flowUrl,
        HttpClient.configurations.v1,
        options
      );

      if (response.ok) {
        alert("Success! Event created via Power Automate.");
      } else {
        const err = await response.text();
        alert("Flow failed: " + err);
      }
    } catch (error) {
      console.error(error);
      alert("Error calling Flow.");
    } finally {
      this.setState({ loading: false });
    }
  }

  public render(): React.ReactElement<IEventsProps> {
    const stackTokens = { childrenGap: 20 };

    return (
      <section style={{ padding: '20px', background: '#f4f4f4', borderRadius: '8px' }}>
        <Stack tokens={stackTokens}>
          <Label style={{ fontSize: '20px', borderBottom: '1px solid #ccc' }}>
            Event Details
          </Label>

          <TextField
            label="Event Title"
            value={this.state.title}
            onChange={(e, val) => this.setState({ title: val || "" })}
          />

          <DateTimePicker
            label="Start"
            value={this.state.startDateTime}
            onChange={(date) => this.setState({ startDateTime: date as Date })}
          />

          <DateTimePicker
            label="End"
            value={this.state.endDateTime}
            onChange={(date) => this.setState({ endDateTime: date as Date })}
          />

          <Dropdown
            label="Calendars"
            placeholder="Select calendars"
            multiSelect
            options={CALENDAR_OPTIONS}
            selectedKeys={this.state.selectedCalendars}
            onChange={(e, o) => this._onDropdownChange(e, o, 'selectedCalendars')}
          />

          <Dropdown
            label="Categories"
            placeholder="Select categories"
            multiSelect
            options={CATEGORY_OPTIONS}
            selectedKeys={this.state.selectedCategories}
            onChange={(e, o) => this._onDropdownChange(e, o, 'selectedCategories')}
          />

          <TextField
            label="Location"
            value={this.state.location}
            onChange={(e, val) => this.setState({ location: val || "" })}
          />

          <TextField
            label="Description"
            multiline
            rows={3}
            value={this.state.description}
            onChange={(e, val) => this.setState({ description: val || "" })}
          />

          <Checkbox
            label="Add Teams link?"
            checked={this.state.addTeamsLink}
            onChange={(e, checked) => this.setState({ addTeamsLink: !!checked })}
          />

          <PrimaryButton
            text={this.state.loading ? "Creating..." : "Create Event"}
            onClick={this._triggerFlow}
            disabled={
              this.state.loading || 
              !this.state.title || 
              !this.state.startDateTime || 
              !this.state.endDateTime || 
              this.state.selectedCalendars.length === 0
            }
          />
        </Stack>
      </section>
    );
  }
}
