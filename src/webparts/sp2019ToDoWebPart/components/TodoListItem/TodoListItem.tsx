import * as React from 'react';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import {
  Button,
  ButtonType,
} from 'office-ui-fabric-react/lib/Button';
import {
  FocusZone,
  FocusZoneDirection
} from 'office-ui-fabric-react/lib/FocusZone';
import { css } from 'office-ui-fabric-react/lib/Utilities';
import styles from './TodoListItem.module.scss';
import ITodoItem from '../../models/ITodoItem';
import ITodoListItemProps from './ITodoListItemProps';
import Todo from '../TodoContainer/TodoContainer';
import { Dropdown } from 'office-ui-fabric-react';
import Swal from 'sweetalert';

export default class TodoListItem extends React.Component<ITodoListItemProps, {}> {
  public  list = [
    {text:'Apparel', key: 'Apparel', value: 'Apparel',listId:'6ac7c5a8-9cff-4d31-bdac-8186a2d198ab'},
    {text:'Consumer', key: 'Consumer', value: 'Consumer',listId:'42108a97-0f75-4c42-8485-5dd9f593d713'},

    {text:'Techtex', key: 'Techtex', value: 'Techtex',listId:'1325b51a-d688-44c6-be0c-71af9a15ddf1'},
   {text:'Automotive', key: 'Automotive', value: 'Automotive',listId:'b5f11715-9d34-416c-9d0c-bd055ee95400'},
 //  {text:'Automotive test', key: 'Automotive test', value: 'Automotive test',listId:'cb8f32e7-38c8-4b14-982c-9db38ec49d1d'}
  
    // {text:'Automotive', key: 'Automotive', value: 'Automotive',listId:'CB8F32E7-38C8-4B14-982C-9DB38EC49D1D'},
    //{text:'Automotive', key: 'Automotive', value: 'Automotive',listId:'3f11210d-0b9d-44e3-8e0d-35067d494b6f'},
  
    // {text:'Automotive', key: 'Customer Documents - Automotive', value: 'Automotive',listId:'4010f6e6-3ee9-4740-a247-20afaa652f1d' },
    // {text:'Consumer', key: 'Customer Documents Consumer', value: 'Consumer',listId:'4010f6e6-3ee9-4740-a247-20afaa652f1d' },
    // { text:'Techtex',key: 'Customer Documents - Techtex', value: 'Techtex' ,listId:'4010f6e6-3ee9-4740-a247-20afaa652f1d'}
   ];

  //  public  list = [
  //   // {text:'Prototyp Automotive Improved', key: 'Prototyp Automotive Improved', value: 'Prototyp Automotive Improved',listId:'4010f6e6-3ee9-4740-a247-20afaa652f1d'},
  //   {text:'Automotive', key: 'Customer Documents - Automotive', value: 'Automotive',listId:'04a2e24d-e933-4634-b633-80f58d86e41e' }
  //   // {text:'Consumer', key: 'Customer Documents Consumer', value: 'Consumer',listId:'4010f6e6-3ee9-4740-a247-20afaa652f1d' },
  //   // { text:'Techtex',key: 'Customer Documents - Techtex', value: 'Techtex' ,listId:'4010f6e6-3ee9-4740-a247-20afaa652f1d'}
  //  ];
  // public  list = [
  //   // {text:'Automotive', key: 'Customer Documents - Automotive', value: 'Automotive',listId:'4010f6e6-3ee9-4740-a247-20afaa652f1d' },
  //   {text:'copied', key: 'copied', value: 'copied',listId:'7dd8c4bb-ba47-4814-949c-6f3e54b99ee8' }
  //   // {text:'Site Collection Documents', key: 'SiteCollectionDocuments', value: 'Site Collection Documents',listId:'cf01a20b-e7cc-4984-a68e-22aa849b6e39' },
  //   // {text:'Customer Documents Automotive', key: 'Customer%20Documents%20%20Automotive', value: 'Customer%20Documents%20%20Automotive',listId:'04a2e24d-e933-4634-b633-80f58d86e41e' },
  //   // { text:'testdoc',key: 'testdoc', value: 'testdoc' ,listId:'a5204a0e-ce52-4374-ae6e-6ef14e8424a3'}
  // ];
  public selectedLibId = "";
  constructor(props: ITodoListItemProps) {
    super(props);

    this.handleDropdownChange = this.handleDropdownChange.bind(this);
    this._handleDeleteClick = this._handleDeleteClick.bind(this);
    this.handleDocEdit= this.handleDocEdit.bind(this);
    this.truncateText = this.truncateText.bind(this)
    TodoListItem.GetIconPath = TodoListItem.GetIconPath.bind(this)
    TodoListItem.getFileExtension = TodoListItem.getFileExtension.bind(this)
  }

  public shouldComponentUpdate(newProps: ITodoListItemProps): boolean {
    return (
      this.props.item !== newProps.item ||
      this.props.isChecked !== newProps.isChecked
    );
  }
  public  minimizeText(text, maxCharacters) {
    if( text!=undefined){
      if (text.length <= maxCharacters) {
        return text;
      }
      const minimizedText = text.slice(0, maxCharacters) + '...';
      return minimizedText;
    }
    return '';
  }

  public render(): JSX.Element {

    const classTodoItem: string = css(
      styles.todoListItem,
      'ms-Grid',
      'ms-u-slideDownIn20'
    );
    let htmlid= "data_"+this.props.item.Id;
    return (
      
      <div
        role='row'
        className={classTodoItem}
        data-is-focusable={true}
      >
        <FocusZone direction={FocusZoneDirection.horizontal}>
        <div className={css(styles.todoListItem)}>
  <div className={css(styles.rowcontainer)}>
  <div className={css(styles.infocontainer)}>
    <img src={TodoListItem.GetIconPath(TodoListItem.getFileExtension(this.props.item.LinkFilename))} alt="File Logo" /> 
    <a href={this.props.item.ServerUrl} className={css(styles.filenamelink)} target="_blank">
      <label className={css(styles.filenamelabel)}>{this.minimizeText(this.props.item.LinkFilename, 35)}</label>
    </a>
    {/* <label className={css(styles.filenamelabel)}>{this.truncateText(this.props.item.LinkFilename, 5)}</label> */}
    <Dropdown
      id={htmlid}
      options={this.list}
      className={css(styles.customdropdown)}
      defaultSelectedKey="Please select.."
      disabled={false}
      multiSelect={false}
      placeHolder="Select division"
      required={true}
      onChanged={this.handleDropdownChange}
    />
  </div>
  <div className={css(styles.buttoncontainer)}>
  <button className={styles.editbutton} onClick={(): void => this._handleEditModalItem(this.props.item)}>Edit</button>
    <button className={styles.editbutton} onClick={(): void => this.handleDocEdit(this.props.item)}>Move</button>
    <button className={styles.deletebutton} onClick={this._handleDeleteClick}>Delete</button>
   
  </div>
</div>
</div>
        </FocusZone>
        
      </div>
    );
  }
  private _handleEditModalItem(item:ITodoItem) {
    window.open(item.DefaultEditUrl, '_blank');
  }

  // private _handleToggleChanged(ev: React.FormEvent<HTMLInputElement>, checked: boolean): void {
  //   const newItem: ITodoItem = update(this.props.item, {
  //     PercentComplete: { $set: this.props.item.PercentComplete >= 1 ? 0 : 1 }
  //   });

  //   this.props.onCompleteListItem(newItem);
  // }

  private _handleDeleteClick(event: React.MouseEvent<HTMLButtonElement>) {
    Swal({
      title: "Are you sure?",
      text: "You are about to delete "+this.props.item.LinkFilename+". This action cannot be undone.",
      icon: "warning",
      buttons: ["Cancel", "Delete"],
      dangerMode: true,
    }).then((confirm) => {
      if (confirm) {
        // Delete action confirmed
        // Perform the delete action here
        this.props.onDeleteListItem(this.props.item);
      }
    });
  }
//#region Event
public handleDropdownChange = (event) => {
  //console.log(event.value);
  this.selectedLibId =  event.listId;
};
public handleDocEdit(item){
  const data = document.getElementById('data_'+item.Id) as HTMLInputElement;
  if(data.id =="data_"+item.Id && this.selectedLibId!= null && this.selectedLibId!=""){

    this.props.onEditListItem(item, this.selectedLibId);
   // this.props.onEditListItem(newItem,"DropLibrary");
  }else{
    Swal({
      icon: 'warning',
      title: 'Validation Error',
      text: 'Please select the Division',
    });
  }

  if(data.disabled){
    data.disabled = false;
  }else{
    data.disabled = true;
  }
  
  // const fileInput = document.getElementById('file-input') as HTMLInputElement;
  // if (fileInput) {
  //   fileInput.value = item.LinkFilename;
  // }
  console.log(this.selectedLibId);
 // this.setState({item:item});
  // this.setState({});
  console.log(item);


  // this.props.
  // if(this.ValidateInputs()){
  //   event.preventDefault();
  //   console.log(this.state);
  //  // this.props.onAddTodoItem(this.state.selectedFiles,this.state.selectedLibrary,this.state.createDocumentSet);
  // }
};
//#endregion
  //#region helper
  public static getFileExtension(title) {
    if(title!=undefined){
      const dotIndex = title.lastIndexOf('.');
      // Check if a dot exists and it is not the last character
      if (dotIndex !== -1 && dotIndex !== title.length - 1) {
        // Extract the substring after the dot
        const extension = title.substring(dotIndex + 1);
        // Return the file extension in lowercase
        return extension.toLowerCase();
      }
    }
    return '';
  }
  public static GetIconPath(extn:string) {
    var imgPath = "";
    switch (extn) {
        case "pptx":
            imgPath = "/_layouts/15/images/icpptx.png";
            break;
        case "mp4":
          imgPath = "/_layouts/15/images/icvidset.gif";
          break;
        case "mov":
          imgPath = "/_layouts/15/images/icvidset.gif";
          break;
          case "msg":
          imgPath = "/_layouts/15/images/icmsg.gif";
          break;
        case "jpg":
          imgPath = "/_layouts/15/images/icjpg.gif";
          break;
        case "ppt":
            imgPath = "/_layouts/15/images/icpptx.png";
            break;
        case "docx":
            imgPath = "/_layouts/15/images/icdocx.png";
            break;
        case "doc":
            imgPath = "/_layouts/15/images/icdocx.png";
            break;
        case "xlsx":
            imgPath = "/_layouts/15/images/icxlsx.png";
            break;
        case "xls":
            imgPath = "/_layouts/15/images/icxlsx.png";
            break;
        case "pdf":
            imgPath = "/_layouts/15/images/icpdf.png";
            break;
        case "xml":
            imgPath = "/_layouts/15/images/icxml.gif";
            break;
        case "xlsm":
            imgPath = "/_layouts/15/images/icxlsm.png";
            break;
        case "csv":
            imgPath = "/_layouts/15/images/icxls.png";
            break;
        case "txt":
            imgPath = "/_layouts/15/images/ictxt.gif";
            break;
        default:
            imgPath = "/_layouts/15/images/icgen.gif";
    }
    return Todo.siteUrl+ imgPath;
  }
  public truncateText(text, limit) {
    const words = text.split(' ');
    if (words.length <= limit) {
      return text;
    } else {
      return words.slice(0, limit).join(' ') + '...';
    }
  }
  showAlert(type, msg, defaultmsg){
    Swal(type, msg, defaultmsg);
  };
  //#endregion
}
