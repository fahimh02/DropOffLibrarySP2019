import * as React from 'react';
import {
  Button,
  ButtonType,
  DefaultButton,
  PrimaryButton
} from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import styles from './TodoForm.module.scss';
import ITodoFormState from './ITodoFormState';
import ITodoFormProps from './ITodoFormProps';
import Swal from 'sweetalert';
export default class TodoForm extends React.Component<ITodoFormProps, ITodoFormState>{

  private _placeHolderText: string = 'Enter your todo';

  constructor(props: ITodoFormProps) {
    super(props);

    this.state = {
      selectedFiles: [],
      selectedLibrary: '',
      createDocumentSet: false,
    };
    this.handleFormSubmit = this.handleFormSubmit.bind(this)
    this.handleFileUpload = this.handleFileUpload.bind(this);
    this.handleClearClick  = this.handleClearClick.bind(this);
  }
  //#region Event
  public handleFormSubmit = (event) => {
    if(this.ValidateInputs()){
      event.preventDefault();
      console.log(this.state);
      this.props.onAddTodoItem(this.state.selectedFiles,this.state.selectedLibrary,this.state.createDocumentSet);
      this.handleClearClick();
    }
  };
  public handleFileUpload = (event) => {
    const files = Array.from(event.target.files);
    this.setState((prevState) => ({
      selectedFiles: [...prevState.selectedFiles, ...files],
    }));
  };
  public handleClearClick = () => {
    const fileInput = document.getElementById('file-input') as HTMLInputElement;
    if (fileInput) {
      fileInput.value = '';
    }
    this.setState({ selectedLibrary: '' });
    this.setState({ selectedFiles: [] });
  };
  //#endregion
  
  ValidateInputs() {
    if(this.state.selectedFiles.length<=0){
      this.showAlert("Warning!", "Please chose a file!","warning");
      return false;
    }
    // if(this.state.selectedLibrary==''){
    //   this.showAlert("Warning!", "Please select the library!","warning");
    //   return false;
    // }
    return true;
  }
  showAlert(type,msg,defaultmsg){
    Swal(type,msg, defaultmsg);
  };

  public render(): JSX.Element {
    return (
      
      <div className={styles.todoForm}>
        <form onSubmit={this.handleFormSubmit} style={{ display: 'flex', flexDirection: 'column',  padding: '20px', borderRadius: '8px' }}>
      <div style={{ marginBottom: '20px' }}>
        <label htmlFor="file" style={{ color: 'navy', fontWeight: 'bold' }}>Upload File:</label>
        <input type="file" id="file-input"  onChange={this.handleFileUpload} />
      </div>
    
      <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: '1rem' }}>
        <PrimaryButton className={styles.uploadbutton} style={{ marginRight: '0.5rem' }}onClick={this.handleFormSubmit} >Upload</PrimaryButton>
        <DefaultButton style={{ backgroundColor: 'lightgray' }} onClick={this.handleClearClick}>Clear</DefaultButton>
      </div>
    </form>
        {/* <TextField
          className={styles.textField}
          value={this.state.inputValue}
          placeholder={this._placeHolderText}
          autoComplete='off'
          onChanged={this._handleInputChange} />
        <div className={styles.addButtonCell}>
          <Button
            className={styles.addButton}
            buttonType={ButtonType.primary}
            ariaLabel='Add a todo task'
            onClick={this._handleAddButtonClick}>
            Add
          </Button>
        </div> */}
      </div>
    );
  }

  // private _handleInputChange(newValue: string) {
  //   this.setState({
  //     inputValue: newValue
  //   });
  // }

  // private _handleAddButtonClick(event?: React.MouseEvent<HTMLButtonElement>) {
  //   this.setState({
  //     inputValue: this._placeHolderText
  //   });
  //   this.props.onAddTodoItem(this.state.inputValue);
  // }
}
