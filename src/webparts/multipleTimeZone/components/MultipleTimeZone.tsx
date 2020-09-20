import * as React from 'react';
import styles from './MultipleTimeZone.module.scss';
import { IMultipleTimeZoneProps } from './IMultipleTimeZoneProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { PrimaryButton, values } from 'office-ui-fabric-react';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { DateTimePicker, DateConvention, TimeConvention, TimeDisplayControlType } from '@pnp/spfx-controls-react/lib/dateTimePicker';
import * as moment from "moment-timezone";
import * as jstz from "jstz";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

const selectorOptions = moment.tz.names()
  .reduce((memo, tz) => {
    memo.push({
      name: tz,
      offset: moment.tz(tz).utcOffset()
    });

    return memo;
  }, [])
  .sort((a, b) => {
    return a.offset - b.offset;
  })
  .reduce((memo, tz) => {
    const timezone = tz.offset ? moment.tz(tz.name).format('Z') : '';

    return memo.concat(`{"key":"${tz.name}", "text":"(GMT${timezone}) ${tz.name}"},`);
  }, "");

const timeZoneOptions = JSON.parse(`[${selectorOptions.substring(0, selectorOptions.length - 1)}]`);

export interface IMultipleTimeZoneState {
  date: Date;
  dateToStore: string;
  selectedTimeZone: IDropdownOption;
  allItems: any;
}

export default class MultipleTimeZone extends React.Component<IMultipleTimeZoneProps, IMultipleTimeZoneState> {

  constructor(props: IMultipleTimeZoneProps) {
    super(props);
    this.state = {
      date: undefined,
      dateToStore: undefined,
      selectedTimeZone: undefined,
      allItems: []
    };
    this.getAllItems();
  }

  private handleChange = (date: Date) => {
    let dateToStore = moment(date).tz(jstz.determine().name()).format('YYYY-MM-DDTHH:mm:ss').concat('Z');
    this.setState({ date: date, dateToStore: dateToStore });
  }

  private onChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    console.log(item);
    this.setState({ selectedTimeZone: item });
  }

  private saveItem = () => {
    sp.web.lists.getByTitle("Learning datetime").items.add({
      Title: "From SPFx",
      SampleDateTime: this.state.dateToStore,
      TimeZone: this.state.selectedTimeZone.key
    }).then(() => {
      this.getAllItems();
    }).catch((err) => {
      console.log(err);
    });
  }

  private getAllItems = () => {
    sp.web.lists.getByTitle("Learning datetime").items.getAll().then((items) => {
      this.setState({ allItems: items });
    }).catch((err) => {
      console.log(err);
    });
  }

  private convertTimeZones = (index: number) => {
    let items = this.state.allItems;
    let convertedTime = moment.tz(items[index].SampleDateTime.slice(0, -1), 'YYYY-MM-DDTHH:mm:ss',
      items[index].TimeZone).tz(jstz.determine().name()).format('YYYY-MM-DDTHH:mm:ss z');
    items[index].convertedTime = convertedTime;
    this.setState({ allItems: items });
  }

  public render(): React.ReactElement<IMultipleTimeZoneProps> {
    return (
      <div className={styles.multipleTimeZone}>
        <div className={styles.container}>
          <div className={styles.row}>
            <Dropdown
              placeholder="Select Timezone"
              label="Select Timezone"
              options={timeZoneOptions}
              selectedKey={this.state.selectedTimeZone ? this.state.selectedTimeZone.key : undefined}
              onChange={this.onChange}
            />
            <DateTimePicker label="Select Date"
              dateConvention={DateConvention.DateTime}
              timeConvention={TimeConvention.Hours24}
              value={this.state.date}
              onChange={this.handleChange}
              timeDisplayControlType={TimeDisplayControlType.Dropdown} />
            <PrimaryButton style={{ marginTop: 10 }} text="Save Item to List" onClick={this.saveItem} />
          </div>
          <div className={styles.row}>
            {this.state.allItems && this.state.allItems.map((item, index) => {
              return (<div>
                <span style={{ paddingRight: 10 }}>{item.Title}</span>
                <span style={{ paddingRight: 10 }}>{item.SampleDateTime}</span>
                <span style={{ paddingRight: 10 }}>{item.TimeZone}</span>
                <a style={{ paddingRight: 10 }} onClick={() => { this.convertTimeZones(index); }}><span>Convert to Local timezone</span></a>
                <span>{item.convertedTime}</span>
              </div>);
            })}
          </div>
        </div>
      </div>
    );
  }
}
