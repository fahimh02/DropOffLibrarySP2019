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
    {text:'Automotive', key: 'Automotive', value: 'Automotive',listId:'b5f11715-9d34-416c-9d0c-bd055ee95400'}];

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
    let htmlid= "data_"+this.props.item.Id;
    let checkboxId= "data_cb_"+this.props.item.Id;
    return (
      <div className={css(styles.documentrow)}>
      <div className={css(styles.documentdetails)}>
          <div className={css(styles.documentname)}>
              <img className={styles.logo} src={TodoListItem.GetIconPath(TodoListItem.getFileExtension(this.props.item.LinkFilename))} alt="Logo"/>
              <a href={this.props.item.ServerUrl} target="_blank">{this.minimizeText(this.props.item.LinkFilename, 32)}</a>
              
          </div>
          <div className="document-caption">{this.props.item.Customer0}</div>
      </div>
      <div className={css(styles.documentactions)}>
      
          
          
         {
         this.props.item.IsMovePermission && 
          <div className={css(styles.moveCSS)} >
            <Dropdown
              id={htmlid}
              options={this.list}
              className={css(styles['dropdown'])}
              defaultSelectedKey="Please select.."
              disabled={false}
              multiSelect={false}
              placeHolder="Select division"
              required={true}
              onChanged={(event, option) => this.handleDropdownChange(event, option, this.props.item)}
            />
            <input
              type="checkbox"
              id={checkboxId}
              className={css(styles.checkbox)}
              onChange={(event) => this.handleCheckboxChange(event, this.props.item)}
              disabled={true}
            />
          <button className={css(styles.button)} onClick={(): void => this._handleEditModalItem(this.props.item)}>Edit</button>
          <button className={css(styles.button)} onClick={(): void => this.handleDocEdit(this.props.item)}>Move</button>
          <button className={css(styles.dltbutton)} onClick={this._handleDeleteClick}>Delete</button></div>
         }
         {
         !this.props.item.IsMovePermission && 
         <div className={css(styles.moveCSS)}>
         <button className={css(styles.button)} onClick={(): void => this._handleEditModalItem(this.props.item)}>Edit</button></div>
         }
      </div>
  </div>

    );
  }
  private _handleEditModalItem(item:ITodoItem) {
    window.open(item.DefaultEditUrl, '_self');
  }
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
  private handleCheckboxChange(event , item:ITodoItem) {

    const isChecked = event.target.checked;
    this.setState({ isChecked }, () => {
      if (isChecked) {
        // Checkbox is checked, perform some action
        //this.props.item.isMoveOutsideFolder = true;
      } else {
        // Checkbox is unchecked, perform some other action
       // this.props.item.isMoveOutsideFolder = false;
      }
    });
  
  }
  // public handleDropdownChange = (event) => {
  //   console.log(event);
  //   this.selectedLibId =  event.listId;
  //   if(this.selectedLibId =='b5f11715-9d34-416c-9d0c-bd055ee95400'){
  //     console.log("AUTOMOTIVE SELECTED")
  //   }else{
  //     console.log("No SELECTED")
  //   }
  //   console.log("selected list id : ", this.selectedLibId);
  //   const data = document.getElementById('data_'+item.Id) as HTMLInputElement;
  //   if(data.id =="data_"+item.Id && this.selectedLibId!= null && this.selectedLibId!=""){
  //     console.log("MOving to the document lib", this.selectedLibId);
  //     this.props.onEditListItem(item, this.selectedLibId);
  //   }
  // };
  public handleDropdownChange(event, option, item) {
    this.selectedLibId =  event.listId;
    const data = document.getElementById('data_cb_'+item.Id) as HTMLInputElement;
    console.log("cbx",data);
    if(data.id =="data_cb_"+item.Id && this.selectedLibId!= null && this.selectedLibId!=""){
      if(this.selectedLibId =='b5f11715-9d34-416c-9d0c-bd055ee95400'){
        data.disabled = false;
      }else{
        data.disabled = true;
      }
    }
  
  }
  public handleDocEdit(item){
    var moveOutsideFolder:boolean = false;
    const data = document.getElementById('data_'+item.Id) as HTMLInputElement;
    const cbdata = document.getElementById('data_cb_'+item.Id) as HTMLInputElement;
    if(!cbdata.disabled){
      if(cbdata.checked){
        moveOutsideFolder= true;
      }else{
        moveOutsideFolder= false;
      }

    }
    
    if(data.id =="data_"+item.Id && this.selectedLibId!= null && this.selectedLibId!=""){
      console.log("MOving to the document lib", this.selectedLibId);
      this.props.onEditListItem(item, this.selectedLibId,moveOutsideFolder);
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
    console.log(this.selectedLibId);
    console.log(item);
    this.resetInputs(item)

  };
  public resetInputs(item){
    const cbdata = document.getElementById('data_cb_'+item.Id) as HTMLInputElement;
    const data = document.getElementById('data_'+item.Id) as HTMLInputElement;
    data.value= null;
    
    cbdata.checked = false;
  }
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
