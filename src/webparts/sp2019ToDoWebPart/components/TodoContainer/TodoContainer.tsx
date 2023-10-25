import * as React from 'react';
import { DisplayMode } from '@microsoft/sp-core-library';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import TodoForm from '../TodoForm/TodoForm';
import styles from './TodoContainer.module.scss';
import ITodoItem from '../../models/ITodoItem';
import ConfigurationView from '../ConfigurationView/ConfigurationView';
import TodoList from '../TodoList/TodoList';
import ITodoContainerProps from './ITodoContainerProps';
import ITodoContainerState from './ITodoContainerState';
import * as update from 'immutability-helper';
import ITodoTaskList from '../../models/ITodoTaskList';
import Swal from 'sweetalert';
import * as ReactDOM from 'react-dom';
import { Dialog } from '@microsoft/sp-dialog';

export default class Todo extends React.Component<ITodoContainerProps, ITodoContainerState> {
  private _showPlaceHolder: boolean = true;
  private _permission:boolean = false;
  public static siteUrl = "";
  constructor(props: ITodoContainerProps) {
    super(props);
    
    Todo.siteUrl = this.props.dataProvider.getSiteUrl();
    if (this.props.dataProvider.selectedList) {
      if (this.props.dataProvider.selectedList.Id !== '0') {
        this._showPlaceHolder = false;
      }
      else if (this.props.dataProvider.selectedList.Id === '0') {
        this._showPlaceHolder = true;
      }
    } else {
      this._showPlaceHolder = true;
    }

    this.state = {
      todoItems: [],
      libraries:[],
      isLoading:false,
      showDialog:false
    };

    this._configureWebPart = this._configureWebPart.bind(this);
    this._createTodoItem = this._createTodoItem.bind(this);
    this._completeTodoItem = this._completeTodoItem.bind(this);
    this._deleteTodoItem = this._deleteTodoItem.bind(this);
    this.showAlert = this.showAlert.bind(this);
  }
  showAlert(type,msg,defaultmsg){
    Swal(type,msg, defaultmsg);
  };
  public componentWillReceiveProps(props: ITodoContainerProps) {
    
    if (this.props.dataProvider.selectedList) {
      if (this.props.dataProvider.selectedList.Id !== '0') {
        this._showPlaceHolder = false;
        this.props.dataProvider.getDocs().then(
          (items: ITodoItem[]) => {
            const newItems = update(this.state.todoItems, { $set: items });
            this.setState({ todoItems: newItems });
          });
      }
      else if (this.props.dataProvider.selectedList.Id === '0') {
        this._showPlaceHolder = true;
      }
    } else {
      this._showPlaceHolder = true;
    }
  }

  
  public componentDidMount() {
    // this.props.dataProvider.getRequiredFields().then((x=> {
    //   console.log("getRequiredFieldsByLibraryName", x.value);
    //   console.log(JSON.stringify(x.value,null,2));
    // }))
    this.setState({isLoading:true});
    let result = this.props.dataProvider.getPermissions().then((res=>{
      if(res==true){
        this._permission = true;
      }else{
        this._permission = false;
      }
    }));
   

    //console.log("getPermissions:",res,this._showPlaceHolder);

    this.props.dataProvider.getDocs().then(
      (items: ITodoItem[]) => {
        console.log("Documents", items);
        const newItems = update(this.state.todoItems, { $set: items });
        this.setState({ todoItems: newItems });
      });
      this.props.dataProvider.getLists().then(
        (items: ITodoTaskList[]) => {
        this.setState({libraries:items,isLoading:false})
      });
    // if (!this._showPlaceHolder) {
    //   this.props.dataProvider.getItems().then(
    //     (items: ITodoItem[]) => {
    //       this.setState({ todoItems: items });
    //     });
    // }
  }

  public render(): JSX.Element {
    const { libraries ,isLoading,showDialog} = this.state;
    console.log("isLoading",isLoading);
    return (
      <Fabric>
          {this._showPlaceHolder && this.props.webPartDisplayMode === DisplayMode.Edit &&
            <ConfigurationView
              icon={'ms-Icon--Edit'}
              iconText='Drop off Library'
              description='Upload files in a drop off library'
              buttonLabel='Configure'
              onConfigure={this._configureWebPart} />
          }
          {this._showPlaceHolder && this.props.webPartDisplayMode === DisplayMode.Read &&
            <ConfigurationView
              icon={'ms-Icon--Edit'}
              iconText='Drop off Library'
              description='Upload files in a drop off library. Edit this web part to start managing drop off library.' />
          }
          {
            // , opacity: isLoading ? 0.5 : 1, pointerEvents: isLoading ? 'none' : 'auto'
            
            
            !this._showPlaceHolder && !showDialog && this._permission && 
            <div className={styles.todo}>
              <div className={styles.topRow}>
                <h2 className={styles.todoHeading}>Drop off library</h2>
              </div>
              {libraries ? (
              <TodoForm list={libraries}  onAddTodoItem={this._createTodoItem} />

            ) : (
              <div>Loading...</div>
            )}
            <div style={{ opacity: isLoading ? 0.4 : 1, pointerEvents: isLoading ? 'none' : 'auto'}}>
              <div className={styles.documentlist}>
              <TodoList items={this.state.todoItems}
              onEditTodoItem={this._completeTodoItem}
              onDeleteTodoItem={this._deleteTodoItem} />
              </div>
            
              </div>
            </div>
          }
            
        
        </Fabric>
     
    );
  }
  private _configureWebPart(): void {
    this.props.configureStartCallback();
  }
  private _createTodoItem(selectedFiles: File[],selectedLibrary:string, createDocumentSet:boolean): Promise<any> {
    this.setState({isLoading:true})
    //console.log(selectedFiles,selectedLibrary,createDocumentSet);
    return this.props.dataProvider.createDoc(selectedFiles).then(
      (items: ITodoItem[]) => {
        let filenamelabel = selectedFiles[selectedFiles.length-1].name; 
        const newItems = update(this.state.todoItems, { $set: items });
        this.setState({ todoItems: newItems,isLoading:false });
        this.showAlert("Success!",filenamelabel+" uploaded successfully in Dropoff Library!","success");
      });
    // return this.props.dataProvider.uploadItems(selectedFiles,selectedLibrary,createDocumentSet).then(
    //   (items: ITodoItem[]) => {
    //     const newItems = update(this.state.todoItems, { $set: items });
    //     this.setState({ todoItems: newItems });
    //   });
  }
  // private _createTodoItem(inputValue: string): Promise<any> {
  //   return this.props.dataProvider.createItem(inputValue).then(
  //     (items: ITodoItem[]) => {
  //       const newItems = update(this.state.todoItems, { $set: items });
  //       this.setState({ todoItems: newItems });
  //     });
  // }
  handleFieldChange(fieldName: string, value: string) {
    // Handle field value changes
   // console.log(`Field '${fieldName}' changed to '${value}'`);
  }
  async saveChanges(itemId: string) {
    // Save the changes to SharePoint
    //console.log('Saving changes...');
    // Perform your logic to save the changes to SharePoint using SPHttpClient or any other appropriate method
  }

  private async _completeTodoItem(todoItem: ITodoItem,libName:string,ismoveoutsideFolder:boolean): Promise<any> {
    try {
      this.setState({isLoading:true})
    const item : ITodoItem = await this.props.dataProvider.getDoc(todoItem.Id.toString());
    console.log("got item:",item);
    todoItem = item;
    let x = await this.props.dataProvider.uploadItems(todoItem,libName,ismoveoutsideFolder);
    if(x.Id!=0){
      return  this.props.dataProvider.deleteDoc(todoItem).then(
          (items: ITodoItem[]) => {
           // console.log("getDocs,",items);
            const newItems = update(this.state.todoItems, { $set: items });

            this.setState({isLoading:false});
            this.setState({ todoItems: newItems });
          });
    }else{
      return this.props.dataProvider.getDocs().then(
        (items: ITodoItem[]) => {
        //  console.log("getDocs,",items);
          const newItems = update(this.state.todoItems, { $set: items });
          this.setState({ todoItems: newItems });
          this.setState({isLoading:false});
        });
    }
      
    } catch (error) {
      console.log(error);
      this.setState({isLoading:false});
    }
    
  }

  private _deleteTodoItem(todoItem: ITodoItem): Promise<any> {
    return this.props.dataProvider.deleteDoc(todoItem).then(
      (items: ITodoItem[]) => {
        const newItems = update(this.state.todoItems, { $set: items });
        this.setState({ todoItems: newItems });
      });
  }

}
