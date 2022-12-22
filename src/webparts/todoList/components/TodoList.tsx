/* eslint-disable @typescript-eslint/no-floating-promises */
import { useEffect, useState } from 'react';
import * as React from 'react';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { SPHttpClient } from "@microsoft/sp-http";

import styles from './TodoList.module.scss';
import { ITodoListProps } from './ITodoListProps';
import { Icon, List, Stack, TextField, Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownStyles } from '@fluentui/react';
import * as strings from 'TodoListWebPartStrings';
import { get, findIndex } from '@microsoft/sp-lodash-subset';
import { SPFI, spfi } from "@pnp/sp";
import { getSP } from '../pnpjsConfig';
import { IItemAddResult, Items } from '@pnp/sp/items';
import useForm from './Form';
import validate from './Validation';
import { ISPListItem, IItemListProps, ITodoFormValues } from './Interface';


const ToDoListName: string = 'To do list'

const _itemStatus = (status: string): string => {
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


const _onRenderListItem = (
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
              className={`${styles.itemStatus} ${_itemStatus(
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


const ItemList = function (props: IItemListProps): React.ReactElement<IItemListProps> {
  const [items, setItems] = useState<ISPListItem[]>([]);
  const { handleSubmit, handleInputChange, inputs, errors } = useForm({ title: "", status: "" }, validate)


  const addTodoItem = async function (values): Promise<void> {
    console.log(values)
    const sp = getSP()
    const list = await sp.web.lists.getByTitle(ToDoListName).select('Title', 'Status');
    const iar: IItemAddResult = await list.items.add({
      Title: values.title,
      Status: values.status
    });
    const item: ISPListItem = { Title: iar.data.Title, Id: iar.data.Id, Status: iar.data.Status };
    setItems([...items, item])
  }

  const _getListData = async function (): Promise<ISPListItem[]> {
    try {
      const response = await props.spHttpClient.get(
        `${props.webUrl}/_api/web/lists/getByTitle('${ToDoListName}')/items?$select=Id,Title,Status`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        const responseText = await response.text();
        throw new Error(responseText);
      }

      const data = await response.json();
      console.log(data)
      setItems(data.value);
      return data.value;
    } catch (error) {
      console.log(error.message)
    }
  }
  useEffect(() => {
    _getListData();
  }, [])

  const statusOptions: IDropdownOption[] = [
    { key: 'Pending', text: 'Pending' },
    { key: 'Completed', text: 'Completed' },
    { key: 'Active', text: 'Active' },
    { key: 'Overdue', text: 'Overdue' }]
  return (
    <>
      <section className="container" >
        <form onSubmit={handleSubmit(addTodoItem)}>
          <Stack horizontal tokens={{ childrenGap: 50 }} verticalAlign="end">

            < TextField label="Title"
              onChange={(e, value) => { e.stopPropagation(); handleInputChange('title', value) }}
              errorMessage={errors && errors.title}
            />

            <Dropdown label="Status" options={statusOptions} styles={{
              dropdown: { width: 300 }
            }}
              errorMessage={errors && errors.status}
              onChange={(e, option) => { console.log(option); handleInputChange('status', option.key.toString()) }}
            />

            <PrimaryButton type="submit">Create</PrimaryButton>
          </Stack>
        </form>
      </section>
      <List items={items} onRenderCell={_onRenderListItem} />
    </>
  );
}

export default class TodoList extends React.Component<ITodoListProps, {}> {

  constructor(props: ITodoListProps) {
    super(props);
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
