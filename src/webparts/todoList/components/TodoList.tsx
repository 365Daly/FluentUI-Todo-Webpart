/* eslint-disable @typescript-eslint/no-floating-promises */
import * as React from 'react';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { SPHttpClient } from "@microsoft/sp-http";

import styles from './TodoList.module.scss';
import { ITodoListProps } from './ITodoListProps';
import { Icon, List, Stack, TextField, Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownStyles } from '@fluentui/react';
import * as strings from 'TodoListWebPartStrings';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPFI, spfi } from "@pnp/sp";
import { getSP } from '../pnpjsConfig';
import { IItemAddResult } from '@pnp/sp/items';


const ToDoListName: string = 'To do list'

export interface ISPListItem {
  Id: string;
  Title: string;
  Status: string;
}


export interface IListItemState {
  items: ISPListItem[];
  newItem: any;
  errorMessage: any;
}

export interface IItemListProps {
  spHttpClient: SPHttpClient;
  webUrl: string;
}

class ItemList extends React.Component<IItemListProps, IListItemState> {
  private _sp: SPFI;

  constructor(props: IItemListProps, state: IListItemState) {
    super(props);

    this.state = {
      items: [],
      newItem: {},
      errorMessage: null
    };
    this.addTodoItem = this.addTodoItem.bind(this);
  }


  private async addTodoItem (): Promise<void> {
    const sp = getSP()
    const list = await sp.web.lists.getByTitle(ToDoListName).select('Title', 'Status');
    const iar: IItemAddResult = await list.items.add({
      Title: this.state.newItem.title,
      Status: this.state.newItem.status
    });
    const newItem: ISPListItem = { Title: iar.data.Title, Id: iar.data.Id, Status: iar.data.Status };
    this.setState({ items: [...this.state.items, newItem] })
  }

  private async _getListData (): Promise<ISPListItem[]> {
    try {
      const response = await this.props.spHttpClient.get(
        `${this.props.webUrl}/_api/web/lists/getByTitle('${ToDoListName}')/items?$select=Id,Title,Status`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        const responseText = await response.text();
        throw new Error(responseText);
      }

      const data = await response.json();
      console.log(data)
      this.setState({ items: data.value });
      return data.value;
    } catch (error) {
      this.setState({ errorMessage: error.message });
    }
  }

  public componentDidMount (): void {
    this._getListData();
  }

  public render (): React.ReactElement<IItemListProps> {
    const txtChange = (e) => {
      console.log(e)
      const newItem = this.state.newItem;
      newItem['title'] = e.target.value;
      this.setState({
        newItem: newItem
      })
    }
    const statusChange = (option) => {
      console.log(option)
      const newItem = this.state.newItem;
      newItem['status'] = option.key;
      this.setState({
        newItem: newItem
      })
    }
    const statusOptions: IDropdownOption[] = [
      { key: 'Pending', text: 'Pending' },
      { key: 'Completed', text: 'Completed' },
      { key: 'Active', text: 'Active' },
      { key: 'Overdue', text: 'Overdue' }]
    return (
      <>
        <section className="container" >
          <Stack horizontal tokens={{ childrenGap: 50 }} verticalAlign="end">
            <TextField label="Title" value={this.state.newItem.title} onChange={(e) => txtChange(e)} />
            <Dropdown label="Status" options={statusOptions} styles={{
              dropdown: { width: 300 }
            }}

              onChange={(e, option) => statusChange(option)}
            />
            <PrimaryButton onClick={this.addTodoItem}>Create</PrimaryButton>
          </Stack>
        </section>
        <List items={this.state.items} onRenderCell={this._onRenderListItem} />
        {this.state.errorMessage && <span>{this.state.errorMessage}</span>}
      </>
    );
  }

  public _itemStatus = (status: string): string => {
    switch (status) {
      case "Pending":
        return styles.itemPending;

      case "Completed":
        return styles.itemCompleted;

      case "Active":
        return styles.itemActive;

      case "Overdue":
        return styles.itemOverdue;

      default:
        return styles.itemStatus;
    }
  };

  public _onRenderListItem = (
    item: ISPListItem,
    index: number
  ): JSX.Element => {
    const removeTodoItem = async function (item: ISPListItem): Promise<void> {
      console.log(item)
      const retVal = confirm("Task will be deleted. Do you want to continue?");
      if (retVal === true) {
        const sp = getSP()
        const list = await sp.web.lists.getByTitle(ToDoListName).select('Title', 'Status');
        await list.items.getById(parseInt(item.Id)).delete()
      }

    }


    return (
      <div key={index} data-is-focusable={true}>
        <ul className={styles.list}>
          <li className={styles.listItem}>
            <span>{item.Title}</span>
            <span style={{ display: 'inline-flex', alignItems: 'center' }}>
              <span
                className={`${styles.itemStatus} ${this._itemStatus(
                  item.Status
                )}`}
              >
                {item.Status}
                <Icon
                  className={styles.itemIcon}
                  iconName={`${item.Status === "Completed" ? "Completed" : null}`}
                />

              </span>

              <DefaultButton
                className={styles.removeBtn}
                text="Remove"
                onClick={e => { removeTodoItem(item) }}
                iconProps={{ iconName: "Delete" }}
              />
            </span>
          </li>
        </ul>
      </div >
    );
  };
}

export default class TodoList extends React.Component<ITodoListProps, {}> {
  private _sp: SPFI;

  constructor(props: ITodoListProps) {
    super(props);

    this._sp = spfi(props.sp);
  }


  public render (): React.ReactElement<ITodoListProps> {
    const {
      // description,
      // isDarkTheme,
      // environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;


    return (
      <>
        <h1 className={styles.headline}>My Task List</h1>
        <section
          className={`${styles.todoList} ${hasTeamsContext ? styles.teams : ""}`}
        >
          <h2>
            {strings.ToDoListHeading}
          </h2>
          <ItemList
            spHttpClient={this.props.spHttpClient}
            webUrl={this.props.websiteUrl}
          />
        </section>

      </>
    );
  }
}
