import * as React from 'react';
import { css, Spinner, TextField, ITextFieldState, ITextFieldProps, Button, ButtonType, ElementType } from 'office-ui-fabric-react';
import {
  IWebPartContext
} from '@microsoft/sp-client-preview';

import { EnvironmentType } from '@microsoft/sp-client-base';

import { ISpfxReactTodoListWebPartProps } from '../ISpfxReactTodoListWebPartProps';

import MockHttpClient from './../Utilities/MockHttpClient';
import SharePointService from './../Services/SharePointService';

export interface ISpfxReactTodoListProps extends ISpfxReactTodoListWebPartProps {
  context: IWebPartContext;
}

export interface ITODOListState
{
  tasks: ITask[];
  loading: boolean;
  newTaskName: string;
}

export interface ITask
{
  title:string;
  completed:boolean;
}

export interface ITODOList
{
  value: ITask[];
}

export default class SpfxReactTodoList extends React.Component<ISpfxReactTodoListProps, ITODOListState> {

  constructor(props: ISpfxReactTodoListProps, state: ITODOListState)
  {
    super(props);

    /* Initialize the state object with our interface definition - empty values and loading = true */
    this.state = {
      tasks: [] as ITask[],
      loading: true,
      newTaskName: ''
    };
  }

  /* React life-cycle: This method will be called on component load */
  public componentDidMount(): void {
      this.getTODOs();
  }

  /* React life-cycle: This method will be called on component changed/update - eg. via PropertyPane panel */
  public componentDidUpdate(prevProps: ISpfxReactTodoListProps, prevState: ITODOListState, prevContext: any): void {
      this.getTODOs();
  }

  private getTODOs(): void {
    if(this.props.context.environment.type == EnvironmentType.Local)
    {
        this.getMockListData().then((response) => {
          this.setState((previousState: ITODOListState, curProps: ISpfxReactTodoListProps): ITODOListState => {
            previousState.loading = false;
            previousState.tasks = response.value;
            return previousState;
          });
        });
    } else
    {
      this.getTODOListData().then((response) => {
          this.setState((previousState: ITODOListState, curProps: ISpfxReactTodoListProps): ITODOListState => {
            previousState.loading = false;
            previousState.tasks = response.value;
            return previousState;
          });
        });
    }
  }

  private getMockListData(): Promise<ITODOList> {
    return MockHttpClient.get(this.props.context.pageContext.web.absoluteUrl).then(() => {
        const listData: ITODOList = {
            value:
            [
                { title: 'Task #1', completed:false},
                { title: 'Task #2', completed:false},
                { title: 'Task #3', completed:false},
                { title: 'Task #4', completed:true},
                { title: 'Task #5', completed:true}
            ]
            };

        return listData;
    }) as Promise<ITODOList>;
  }

  private getTODOListData(): Promise<ITODOList> {
    return SharePointService.getTodos(this.props.context.pageContext.web.absoluteUrl, this.props.listName, this.props.hideFinishedTasks).then((todos) => {
      return todos;
    }) as Promise<ITODOList>;
  }

  private _textInput : TextField;

  public render(): JSX.Element {
    const loading: JSX.Element = this.state.loading ? <div style={{margin: '0 auto'}}><Spinner label={'Loading...'} /></div> : <div/>;
    const results: JSX.Element[] = this.state.tasks.map((res: ITask, i: number) => {
      return(
        <div>{res.title}</div>
      );
    });

    return (
      <div>
      <TextField label="Add a new task:" defaultValue="" ref={component => this._textInput = component} onChanged={this.onInputChanged.bind(this)} />
       <Button elementType={ElementType.button} buttonType={ButtonType.primary} onClick={this.onAddItemClicked.bind(this)}>Add</Button>

        {loading}
        {results}
      </div>
    );
  }

  private onInputChanged(newValue) : void
  {
    this.setState((previousState: ITODOListState, curProps: ISpfxReactTodoListProps): ITODOListState => {
      previousState.newTaskName = newValue;
      return previousState;
    });
  }

  private onAddItemClicked() :void {

    SharePointService.addTodo(this.props.context.pageContext.web.absoluteUrl, this.props.listName, this.state.newTaskName).then(() =>
    {
      SharePointService.getTodos(this.props.context.pageContext.web.absoluteUrl, this.props.listName, this.props.hideFinishedTasks);
    });

    this.setState((previousState: ITODOListState, curProps: ISpfxReactTodoListProps): ITODOListState => {
      previousState.newTaskName = '';
      return previousState;
    });

      this._textInput.setState((previousState: ITextFieldState, curProps: ITextFieldProps): ITextFieldState => {
      previousState.value = '';
      return previousState;
    });
  }
}
