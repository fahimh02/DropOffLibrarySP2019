import { IWebPartContext } from '@microsoft/sp-webpart-base';
import ITodoItem from '../models/ITodoItem';
import ITodoTaskList from '../models/ITodoTaskList';

interface ITodoDataProvider {

  selectedList: ITodoTaskList;

  webPartContext: IWebPartContext;

  getTaskLists(): Promise<ITodoTaskList[]>;

  getItems(): Promise<ITodoItem[]>;

  createItem(title: string): Promise<ITodoItem[]>;

  updateItem(itemUpdated: ITodoItem): Promise<ITodoItem[]>;

  deleteItem(itemDeleted: ITodoItem): Promise<ITodoItem[]>;
 // #region custom
  // uploadItems(itemDeleted: ITodoItem,selectedLibrary:string,createDocumentSet:boolean): Promise<ITodoItem[]>;
  uploadItems(itemDeleted: ITodoItem,selectedLibrary:string,createDocumentSet:boolean):Promise<ITodoItem>;
  getSiteUrl():string;
  getLists():Promise<any>;
  getDocs(): Promise<ITodoItem[]>;
  getDoc(itemId:string): Promise<ITodoItem>;
  createDoc(document: File[]): Promise<ITodoItem[]>;
  deleteDoc(itemDeleted: ITodoItem): Promise<ITodoItem[]>;
  getPermissions():Promise<boolean>;

  //#endregion
}

export default ITodoDataProvider;