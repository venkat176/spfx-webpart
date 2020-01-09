import * as React from 'react';
import styles from './Helloworld.module.scss';
import { IHelloworldProps } from './IHelloworldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SpinButton } from 'office-ui-fabric-react/lib/SpinButton';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { TextField, MaskedTextField } from 'office-ui-fabric-react/lib/TextField';
import { IStackTokens, Stack } from 'office-ui-fabric-react/lib/Stack';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Checkbox, ICheckboxProps } from 'office-ui-fabric-react/lib/Checkbox';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { IStackProps } from 'office-ui-fabric-react/lib/Stack';
import { BaseComponent, assign } from 'office-ui-fabric-react/lib/Utilities';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { IPersonaProps, Persona } from 'office-ui-fabric-react/lib/Persona';
import pnp, { sp, Item, ItemAddResult, ItemUpdateResult,Web } from "sp-pnp-js";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

import {
  CompactPeoplePicker,
  IBasePickerSuggestionsProps,
  IBasePicker,
  ListPeoplePicker,
  NormalPeoplePicker,
  ValidationState
} from 'office-ui-fabric-react/lib/Pickers';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { CommandBar, ICommandBarItemProps } from 'office-ui-fabric-react/lib/CommandBar';
import { IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { Label } from 'office-ui-fabric-react/lib/Label';

// import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";


const overflowProps: IButtonProps = { ariaLabel: 'More commands' };

export interface IButtonExampleProps {
  // These are set based on the toggles shown above the examples (not needed in real code)
  disabled?: boolean;
  checked?: boolean;
}

const _items: ICommandBarItemProps[] = [
  {
    key: 'newItem',
    text: 'New',
    cacheKey: 'myCacheKey', // changing this key will invalidate this item's cache
    iconProps: { iconName: 'Add' },
    subMenuProps: {
      items: [
        {
          key: 'emailMessage',
          text: 'Email message',
          iconProps: { iconName: 'Mail' },
          ['data-automation-id']: 'newEmailButton' // optional
        },
        {
          key: 'calendarEvent',
          text: 'Calendar event',
          iconProps: { iconName: 'Calendar' }
        }
      ]
    }
  },
  {
    key: 'upload',
    text: 'Upload',
    iconProps: { iconName: 'Upload' },
    href: 'https://dev.office.com/fabric'
  },
  {
    key: 'share',
    text: 'Share',
    iconProps: { iconName: 'Share' },
    onClick: () => console.log('Share')
  },
  {
    key: 'download',
    text: 'Download',
    iconProps: { iconName: 'Download' },
    onClick: () => console.log('Download')
  }
];

const _overflowItems: ICommandBarItemProps[] = [
  { key: 'move', text: 'Move to...', onClick: () => console.log('Move to'), iconProps: { iconName: 'MoveToFolder' } },
  { key: 'copy', text: 'Copy to...', onClick: () => console.log('Copy to'), iconProps: { iconName: 'Copy' } },
  { key: 'rename', text: 'Rename...', onClick: () => console.log('Rename'), iconProps: { iconName: 'Edit' } }
];

const _farItems: ICommandBarItemProps[] = [
  {
    key: 'tile',
    text: 'Grid view',
    // This needs an ariaLabel since it's icon-only
    ariaLabel: 'Grid view',
    iconOnly: true,
    iconProps: { iconName: 'Tiles' },
    onClick: () => console.log('Tiles')
  },
  {
    key: 'info',
    text: 'Info',
    // This needs an ariaLabel since it's icon-only
    ariaLabel: 'Info',
    iconOnly: true,
    iconProps: { iconName: 'Info' },
    onClick: () => console.log('Info')
  }
];


// export interface IPeoplePickerExampleState {
//   currentPicker?: number | string;
//   delayResults?: boolean;
//   peopleList: IPersonaProps[];
//   mostRecentlyUsed: IPersonaProps[];
//   currentSelectedItems?: IPersonaProps[];
//   isPickerDisabled?: boolean;
// }

const suggestionProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: 'Suggested People',
  mostRecentlyUsedHeaderText: 'Suggested Contacts',
  noResultsFoundText: 'No results found',
  loadingText: 'Loading',
  showRemoveButtons: true,
  suggestionsAvailableAlertText: 'People Picker Suggestions available',
  suggestionsContainerAriaLabel: 'Suggested contacts'
};

const limitedSearchAdditionalProps: IBasePickerSuggestionsProps = {
  searchForMoreText: 'Load all Results',
  resultsMaximumNumber: 10,
  searchingText: 'Searching...'
};

const limitedSearchSuggestionProps: IBasePickerSuggestionsProps = assign(limitedSearchAdditionalProps, suggestionProps);

export interface ITextFieldMultilineExampleState {
  multiline: boolean;
}


export interface Helloworldstate {
  age: number;
  firstDayOfWeek: DayOfWeek;
  selectedItem?: { key: string | number | undefined };
  currentPicker?: number | string;
  delayResults?: boolean;
  currentSelectedItems?: IPersonaProps[];
  isPickerDisabled?: boolean;
  users: any[];
}

const inputProps: ICheckboxProps['inputProps'] = {
  onFocus: () => console.log('Checkbox is focused'),
  onBlur: () => console.log('Checkbox is blurred')
};

const DayPickerStrings: IDatePickerStrings = {
  months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],

  shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],

  days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],

  shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],

  goToToday: 'Go to today',
  prevMonthAriaLabel: 'Go to previous month',
  nextMonthAriaLabel: 'Go to next month',
  prevYearAriaLabel: 'Go to previous year',
  nextYearAriaLabel: 'Go to next year',
  closeButtonAriaLabel: 'Close date picker'
};

const options: IDropdownOption[] = [
  { key: 'fruitsHeader', text: 'Fruits', itemType: DropdownMenuItemType.Header },
  { key: 'apple', text: 'Apple' },
  { key: 'banana', text: 'Banana' },
  { key: 'orange', text: 'Orange', disabled: true },
  { key: 'grape', text: 'Grape' },
  { key: 'divider_1', text: '-', itemType: DropdownMenuItemType.Divider },
  { key: 'vegetablesHeader', text: 'Vegetables', itemType: DropdownMenuItemType.Header },
  { key: 'broccoli', text: 'Broccoli' },
  { key: 'carrot', text: 'Carrot' },
  { key: 'lettuce', text: 'Lettuce' }
];

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 }
};

const controlClass = mergeStyleSets({
  control: {
    margin: '0 0 15px 0',
    maxWidth: '300px'
  }
});

export default class Helloworld extends React.Component<IHelloworldProps, Helloworldstate, ITextFieldMultilineExampleState> {

  private _picker = React.createRef<IBasePicker<IPersonaProps>>();

  constructor(props) {
    super(props);
    this.state = {
      age: 5,
      firstDayOfWeek: DayOfWeek.Monday,
      selectedItem: undefined,
      currentPicker: 1,
      delayResults: false,
      currentSelectedItems: [],
      isPickerDisabled: false,
      users: []
    }
  }

  public render(): React.ReactElement<IHelloworldProps> {

    const { firstDayOfWeek } = this.state;

    const { selectedItem } = this.state;

    // const columnProps: Partial<IStackProps> = {
    //   tokens: { childrenGap: 15 },
    //   styles: { root: { width: 300 } }
    // };

    const suffix = 'years';

    return (
      <div className={styles.helloworld}>
        <div className={styles.container}>
          <div>
            <CommandBar
              items={_items}
              overflowItems={_overflowItems}
              overflowButtonProps={overflowProps}
              farItems={_farItems}
              ariaLabel="Use left and right arrow keys to navigate between commands"
            />
          </div>
          <div className={styles.formcontainer}>
            {/* <div className={style.row}>
            <div className={style.column}> */}
            {/* <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>{this.props.name} */}
            <div className={styles.rowHead}>
              <p className={styles.description}>{escape(this.props.description)}</p>
              <p >{this.props.name}</p>
              <p>{this.props.state}</p>
            </div>
            <div className={styles.rows}>
              <div className={styles.columns}>
                <TextField label="Name" placeholder="Please enter name here" />
              </div>
              <div className={styles.columns}>
                <Label>Age:</Label>
                <SpinButton
                  // label={''}
                  min={10}
                  max={25}
                  value={'10' + suffix}
                  onValidate={(value: string) => {
                    value = this._removeSuffix(value, suffix);
                    if (Number(value) > 100 || Number(value) < 0 || value.trim().length === 0 || isNaN(+value)) {
                      return '0' + suffix;
                    }

                    return String(value) + suffix;
                  }}
                  onIncrement={(value: string) => {
                    value = this._removeSuffix(value, suffix);
                    if (Number(value) + 1 > 100) {
                      return String(+value) + suffix;
                    } else {
                      return String(+value + 1) + suffix;
                    }
                  }}
                  onDecrement={(value: string) => {
                    value = this._removeSuffix(value, suffix);
                    if (Number(value) - 1 < 0) {
                      return String(+value) + suffix;
                    } else {
                      return String(+value - 1) + suffix;
                    }
                  }}
                  onFocus={() => console.log('onFocus called')}
                  onBlur={() => console.log('onBlur called')}
                  incrementButtonAriaLabel={'Increase value by 2'}
                  decrementButtonAriaLabel={'Decrease value by 2'}
                />
              </div>
            </div>
            <div className={styles.rows}>
              <div className={styles.columns}>
                <Label required={true}>Gender:</Label>
                <ChoiceGroup
                  className="defaultChoiceGroup"
                  defaultSelectedKey="A"
                  options={[
                    {
                      key: 'A',
                      text: 'Male'
                    },
                    {
                      key: 'B',
                      text: 'Female'
                    },
                    {
                      key: 'C',
                      text: 'Others',
                      disabled: false
                    }
                  ]}
                  onChange={onChange}
                //ariaLabelledBy={labelId}
                />
              </div>
              <div className={styles.columns}>
                <DatePicker
                  className={controlClass.control}
                  label="DOB"
                  isRequired={false}
                  firstDayOfWeek={firstDayOfWeek}
                  strings={DayPickerStrings}
                  placeholder="Select a date..."
                  ariaLabel="Select a date"
                />
              </div>
            </div>
            <div className={styles.rows}>
              <div className={styles.columns}>
                <Dropdown
                  label="State"
                  selectedKey={selectedItem ? selectedItem.key : undefined}
                  onChange={this._onChange}
                  placeholder="Select an option"
                  options={[
                    { key: 'fruitsHeader', text: 'South', itemType: DropdownMenuItemType.Header },
                    { key: 'andhra', text: 'Andhra' },
                    { key: 'telangana', text: 'Telangana' },
                    { key: 'karnataka', text: 'karnataka', disabled: true },
                    { key: 'tamilnadu', text: 'Tamilnadu' },
                    { key: 'divider_1', text: '-', itemType: DropdownMenuItemType.Divider },
                    { key: 'vegetablesHeader', text: 'North', itemType: DropdownMenuItemType.Header },
                    { key: 'gujarat', text: 'Gujarat' },
                    { key: 'haryana', text: 'Haryana' },
                    { key: 'punjab', text: 'Punjab' }
                  ]}
                />
              </div>
              <div className={styles.columns}>
                <PeoplePicker
                  context={this.props.context}
                  titleText="People Picker"
                  personSelectionLimit={3}
                  groupName={""} // Leave this blank in case you want to filter from all users    
                  showtooltip={true}
                  isRequired={true}
                  disabled={false}
                  ensureUser={true}
                  selectedItems={this._getPeoplePickerItems}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000} />  
                   </div>
            </div>
            <div className={styles.rows}><TextField label="Address" multiline rows={3} /></div>
            <div className={styles.rows}><Checkbox label="Agreed" onChange={_onChange} /></div>
            <div className={styles.rows}><PrimaryButton text="save" onClick={_alertClicked} allowDisabledFocus /></div>
            <div>
              {/* <Checkbox label="Controlled checkbox" checked={isChecked} onChange={onChange} /> */}
            </div>
          </div>
        </div>
      </div>
    );
  }
  private _onChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    console.log(`Selection change: ${item.text} ${item.selected ? 'selected' : 'unselected'}`);
    this.setState({ selectedItem: item });
  };
  private _hasSuffix(value: string, suffix: string): Boolean {
    const subString = value.substr(value.length - suffix.length);
    return subString === suffix;
  }
  private _getPeoplePickerItems(items: any[]) {
    console.log('Items:', items);
    this.setState({ users: items });
  }

  private _onFilterChanged = (
    filterText: string,
    currentPersonas: IPersonaProps[],
    limitResults?: number
  ): IPersonaProps[] | Promise<IPersonaProps[]> => {
    if (filterText) {
      //let filteredPersonas: IPersonaProps[] = this._filterPersonasByText(filterText);

      //filteredPersonas = this._removeDuplicates(filteredPersonas, currentPersonas);
      //filteredPersonas = limitResults ? filteredPersonas.splice(0, limitResults) : filteredPersonas;
      //return this._filterPromise(filteredPersonas);
    } else {
      return [];
    }
  };

    //   @autobind
    // private addSelectedUsers(): void {
    //     sp.web.lists.getByTitle("SPFx Users").items.add({
    //         Title: getGUID(),
    //         Users: {
    //             results: this.state.addUsers
    //         }
    //     }).then(i => {
    //         console.log(i);
    //     });
    // }

  private AddEventListeners(): void {
    document.getElementById('AddItem').addEventListener('click', () => this.AddItem());
    // document.getElementById('UpdateItem').addEventListener('click',()=>this.UpdateItem());
    // document.getElementById('DeleteItem').addEventListener('click',()=>this.DeleteItem());
  }

  private AddItem(): void {
    pnp.sp.web.lists.getByTitle('kendo Grid').items.add({
      Title: "Hulk"
      // Experience : document.getElementById('Experience')["value"],
      // Location:document.getElementById('Location')["value"]
    });
    alert(" Added !");
  }

  private _removeSuffix(value: string, suffix: string): string {
    if (!this._hasSuffix(value, suffix)) {
      return value;
    }

    return value.substr(0, value.length - suffix.length);
  }
}
function _onChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean) {
  console.log(`The option has been changed to ${isChecked}.`);
}
function onChange(ev: React.FormEvent<HTMLInputElement>, option: IChoiceGroupOption): void {
  console.dir(option);
}
function _alertClicked(): void {
  pnp.sp.web.lists.getByTitle('kendo Grid').items.add({
    Title: "Hulk"
  });
  alert(" Added !");
}

