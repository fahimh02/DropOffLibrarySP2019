import {
  SPHttpClient,
  SPHttpClientBatch,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import ITodoDataProvider from '../dataProviders/ITodoDataProvider';
import ITodoItem from '../models/ITodoItem';
import ITodoTaskList from '../models/ITodoTaskList';
import Swal from 'sweetalert';
import * as JQuery from 'jquery';
import { Web ,sp,Folder ,FolderAddResult} from "@pnp/sp";
import { Item, ItemAddResult, ItemUpdateResult } from '@pnp/sp';
import { resultContent } from 'office-ui-fabric-react/lib-es2015/components/pickers/PeoplePicker/PeoplePicker.scss';
import { filter } from 'lodash';

//import { Web ,sp,Folder} from '@pnp/sp/presets/all';
export default class SharePointDataProvider implements ITodoDataProvider {
  private _selectedList: ITodoTaskList;
  private _taskLists: ITodoTaskList[];
  private static _taskLists: ITodoTaskList[];
  private _listsUrl: string;
  private _listItemsUrl: string;
  private _webPartContext: IWebPartContext;
  private _docItemsUrl: string;
  private documentSetContentTypeName = null;
  private documentContentTypeName = null;
  public consumerListUr;
  public apparalListUrl;
  public techtexListUrl;
  public automotiveListUrl;

  public set selectedList(value: ITodoTaskList) {
    this._selectedList = value;
    this._listItemsUrl = `${this._listsUrl}(guid'${value.Id}')/items`;
    this._docItemsUrl = `${this._listsUrl}(guid'${value.Id}')`;
    //this.documentSetContentTypeName = "document set";
    this.documentSetContentTypeName = "doc set automotive";
    this.documentContentTypeName = "customer document automotive";
    this.consumerListUr = "/contract_mgmt/cd/Customer%20Documents%20Consumer";
    this.apparalListUrl = "/contract_mgmt/cd/Customer%20Documents%20Apparel";
    this.techtexListUrl = "/contract_mgmt/cd/Customer%20Documents%20%20Techtex";
    this.automotiveListUrl = "/contract_mgmt/cd/Customer%20Documents%20%20Automotive"; //prod
    sp.setup({
      sp: {
        baseUrl: "http://intranet.amann.com/contract_mgmt/cd"
      }
    });

  }
  public async getFolderDetail(listUrl:string,targetFolderName:string,contentTypeId:string) {
    sp.setup({
      sp: {
        baseUrl: "http://intranet.amann.com/contract_mgmt/cd"
      }
    });
    try {
      let etfn = await sp.web.getList(listUrl).getListItemEntityTypeFullName();
      console.log("eftn",etfn);
      // first create folder
      let far: FolderAddResult = await sp.web.folders.add(listUrl + '/' + targetFolderName)
      console.log("far",far);
      let fData: any = await sp.web.getFolderById(far.data.UniqueId).select('ID').listItemAllFields.get();            
      console.log("fData",fData);
      return fData;
    } catch (error) {
      console.error('An error occurred:', error);
      return null;
    }
  }
  public async createFolder(libId:string, documentSetContentTypeId:string, folderName:string){
    try {
      const libraryExists = await sp.web.lists.getById(libId).get();
      if (libraryExists) {
         return await sp.web.lists.getById(libId).rootFolder.folders.get().then((folders) => {
          const folderExists = folders.some((folder) => folder.Name === folderName);
          if (folderExists) {
            console.log(`Folder '${folderName}' already exists.`);
            return null;
          } else {

            // The folder doesn't exist, so create it
             return sp.web.lists.getById(libId).rootFolder.folders.add(folderName).then(async(currentFolder) => {
              console.log(`Folder '${folderName}' created successfully in '${libId}'`);
              return await this.getFolderDetail(this.automotiveListUrl, folderName, documentSetContentTypeId);
            });
          }
        }).catch((error) => {
          // Handle errors
          console.error('An error occurred:', error);
        });
      } else {
        console.log(`Document library '${libId}' not found.`);
        return null;
      }
      
    } catch (error) {
      
    }
    return null;
  }
  public async uploadItems(item:ITodoItem, selectedLibrary:string, isMoveOutsideFolder:boolean): Promise<ITodoItem>{
    //var IsW = await this.isOwner(this._webPartContext);


    if(selectedLibrary.toLowerCase().includes("automotive") || selectedLibrary.toLowerCase().includes("b5f11715-9d34-416c-9d0c-bd055ee95400")){
      if(isMoveOutsideFolder){
        return await this._uploadItemInDoclib(item,selectedLibrary)
        .then((itemresponse =>{
          return itemresponse;
          }
        ))
      }else{
        return await this._uploadItemsInDocSet(item,selectedLibrary)
        .then((itemresponse =>{
          return itemresponse;
          }
        ));
      }
    }else{
      return await this._uploadItemInDoclib(item,selectedLibrary)
    .then((itemresponse =>{
      return itemresponse;
      }
    ))
    }
  }
  public async _uploadItemsInDocSet(item:ITodoItem, selectedLibrary:string):Promise<ITodoItem> {
    let statusReq:boolean = false;
    try {
          const documentContent:any  = await this.getFileFormServer(item);
         if(documentContent !== undefined)
         {
           const filteredList =SharePointDataProvider._taskLists.filter(item => item.Id ===selectedLibrary)[0];
           console.log("filteredList: ",filteredList)
           console.log("selectedLibrary: ",selectedLibrary);
           let selectedLibraryTitle = filteredList.Title;
           let selectedListGuid = filteredList.Id;

           //this.convertText( filteredList.Title);

           var internalName= SharePointDataProvider.convertEscapedString(filteredList.EntityTypeName);
           if(internalName =="Prototyp_x0020_Automotive_x0020_Improved"){
           internalName = 'PrototypAutomotiveImproved';
           }
         
           if(filteredList.EntityTypeName == "Customer_x0020_Documents_x0020__x0020_Automotive"){
            internalName = "CustomerDocumentsAutomotive";
           }

           console.log(filteredList, this._listsUrl);
           var cttype  = await this.getAvaiableContentTypesByDocRef(selectedListGuid);
           
           var docsetContentTypeId='' ;
           var documentContentTypeId='' ;
           console.log(cttype,selectedLibraryTitle)
           if(cttype!= undefined &&  cttype.value.length>0)
           {
             let contentTypes = cttype.value;
             const docsetContentType =contentTypes.filter(item => item.Name.toLowerCase().includes(this.documentSetContentTypeName))[0];
             const documentContentType = contentTypes.filter(item => item.Name.toLowerCase().includes(this.documentContentTypeName))[0];
             console.log("docsetContentType",docsetContentType);
             console.log("documentContentType",documentContentType);
             docsetContentTypeId = docsetContentType["Id"]["StringValue"];
             documentContentTypeId = documentContentType["Id"]["StringValue"];
             let documentSetname = this.removeFileExtension(item.LinkFilename);
             const libId ='B5F11715-9D34-416C-9D0C-BD055EE95400';
             
            
             //let test = await this.moveFileByPath("","",item.LinkFilename,true,true);
            
             var docset = await this.createFolder(libId, docsetContentTypeId, documentSetname);
             console.log("dcreate folder rtn",docset);
             //var docset =  null;//await this.getCretedDocumentSetIdV2(item,filteredList,docsetContentTypeId,documentSetname,internalName);
          //  var testdocset =  this.createFolder("Customer%20Documents%20%20Automotive","WOWJQ",docsetContentTypeId);
           // console.log("createFolder return :",testdocset);
            
            //  if(docset!=undefined && docset!=null && docset.d !=undefined && docset.d.Id!=undefined ){
            if( docset!= null){
              //let docsetId = docset.d.Id;
              let docsetId =  docset.ID;  
              let docfile = documentContent as ArrayBuffer;
              let file:File = documentContent;
              const selectedFiles:File[]= [];
              selectedFiles.push(file);
              let ext = this.getFileExtension(item.LinkFilename);
              let fileName = documentSetname;
          
              var modifiedStr = fileName.replace(/\s+\./g, '.');// additional space
              var modifiedStr = modifiedStr.replace(/-/g, '');// remove hyfen
              fileName = this.minimizeText(modifiedStr,30)+'.'+ext;
          
              let responseupload = await this.finalUploadDocset(item,selectedLibraryTitle,selectedFiles,docsetId,documentContentTypeId);
              // let responseupload = undefined;
              console.log("return of finalUploadDocset :",responseupload)
              var updateDocSet = await this.updateDocumentSetById(item,filteredList,docsetContentTypeId,documentSetname,internalName);
              console.log("return of updateDocumentSetById :",updateDocSet)
              if(updateDocSet !=undefined){
              
                this.showAlert("Success!", file.name+" moved sucessfully!","success");
                statusReq = true;
                return item;
              }else if(docset==undefined || docset ==null){
                this.showAlert("Error!", "There is a problem moving the file! : ","error");
                item.Id=0;
                return item;
              }
              else{
                this.showAlert("Error!", "There is a problem moving the file! : ","error");
                item.Id=0;
                return item;
              }
             
            }
             else{
    
               this.showAlert("Error!", "Same name might exists ins destination library!","error");
               item.Id=0;
               return item;
             }
           }
           else{
             this.showAlert("Error!", "no content type matched","error");
             console.log("no file returned from server");
             item.Id=0;
              return item;
           }
       }
      
    } catch (error) {
      console.log("Something went wrong!",error);
      this.showAlert("Error!", "There is a problem creating the Folder! Name might exists.","error");
       item.Id=0;
       return item;
    }
    finally {
      // Code that should always run

      
      if(!statusReq){
        item.Id=0;
      }
      return item;
    }
  

  }
  public async _uploadItemInDoclib(item:ITodoItem, selectedLibrary:string):Promise<ITodoItem> {
    let statusReq:boolean = false;
    try {

          //let documentSetname = this.removeFileExtension(item.LinkFilename);
          
          const documentContent:any  = await this.getFileFormServer(item);
          // const filetoupload = await this.readFile(documentContent)
         if(documentContent !== undefined)
         {
           const filteredList =SharePointDataProvider._taskLists.filter(item => item.Id ===selectedLibrary)[0];
           let selectedListGuid = filteredList.Id;
           selectedLibrary = filteredList.Title;
           this.convertText( filteredList.Title);
           var cttype  = await this.getAvaiableContentTypesByDocRef(selectedListGuid);

           var documentContentTypeId='' ;
           if(cttype!= undefined &&  cttype.value.length>0)
           {
             let contentTypes = cttype.value;
             this.documentContentTypeName = "customer document";
             const documentContentType = contentTypes.filter(item => item.Name.toLowerCase().includes(this.documentContentTypeName))[0];
             documentContentTypeId = documentContentType["Id"]["StringValue"];
             let documentSetname = this.removeFileExtension(item.LinkFilename);
             let docfile = documentContent as ArrayBuffer;
             let file:File = documentContent;
             const selectedFiles:File[]= [];
             selectedFiles.push(file);
             let ext = this.getFileExtension(item.LinkFilename);
             let fileName = documentSetname;
             var modifiedStr = fileName.replace(/\s+\./g, '.');// additional space
             var modifiedStr = modifiedStr.replace(/-/g, '');// remove hyfen
             fileName = this.minimizeText(modifiedStr,30)+'.'+ext;

            let responseupload = await this.finalUploadV2(item,filteredList.Title,filteredList.Id,selectedFiles,documentContentTypeId);
            if(responseupload !=undefined){
              let uploadedItem =  await this.getItemByServerRelativeUrl(responseupload.ServerRelativeUrl);
              if(uploadedItem!= undefined && uploadedItem.value!= undefined && uploadedItem!= null){
                let id = uploadedItem.value;
                let updateupload = await this.updateUpload(id,item,filteredList.Title,filteredList.Id,selectedFiles);
                this.showAlert("Success!", file.name+" moved sucessfully!","success");
                statusReq = true;
                return item;
              }else{
                this.showAlert("Error!", "There is a problem moving the file! : ","error");
                item.Id=0;
                return item;
              }
              
            }
            else{
              this.showAlert("Error!", "There is a problem moving the file! : ","error");
              item.Id=0;
              return item;
            }
             
            
           }
           else{
             this.showAlert("Error!", "no content type matched","error");
             console.log("no file returned from server");
             item.Id=0;
              return item;
           }
       }
      
    } catch (error) {
      console.log("Something went wrong!",error);
      this.showAlert("Error!", "There is a problem creating the document set! Name might exists.","error");
       item.Id=0;
       return item;
    }
    finally {
      // Code that should always run
      console.log('Finally block executed');
    
      
      if(!statusReq){
        item.Id=0;
      }
      return item;
    }
  

  }
  public async getFileFormServer(item:ITodoItem) {
    try {
      let file:File;
      //let currentWebUrl = this._webPartContext.pageContext.site.absoluteUrl;  
      const response = await this._webPartContext.spHttpClient.get(item.ServerUrl, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) =>{
          if (response.ok){
            return response.arrayBuffer();
          }}).then(document =>{
            file = new File([document], item.LinkFilename, { type: this.getFileTypeByExtension(this.removeFileExtension(item.LinkFilename)) });
            // console.log("converted to file:",file);
            return file ;    
          });
          return file; 

    }catch (error) {
      console.log(error);
          return [];
    }
  }
  public async getAvaiableContentTypesByDocRef(documentLibraryGuid){
    try {
      let reqURL = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists/GetById('${documentLibraryGuid}')/contenttypes?$select=Id,Name`;
     // let reqURL = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${documentLibraryName}')/contenttypes?$select=Id,Name`;
      const response = await this._webPartContext.spHttpClient.get(reqURL, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) =>{
        if (response.ok){
          return response.json();
        }});
      return response;
    } catch (error) {
      console.log(error);
          return [];
    }
  }
  public async getContentFieldsById(contentTypeId: string): Promise<any[]> {
    const endpoint = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/ContentTypes('${contentTypeId}')/Fields`;

    return await this._webPartContext.spHttpClient.get(endpoint, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        } else {
          throw new Error(`Error retrieving content type fields. Status: ${response.status} - ${response.statusText}`);
        }
      })
      .then((data: any) => {
        return data.value;
      })
      .catch((error: any) => {
        console.error('Error retrieving content type fields:', error);
        return [];
      });
  }
  public async updateDocumentSetByIdV2(docset,docItem:ITodoItem,filteredList:ITodoTaskList,contentTypeId,documentSetname,internalName){
    let docsetId = docset.d.Id;
    const itemPayload = {
      '__metadata': {
        'type': docset.d['__metadata']['type'],

      },
      'Title': documentSetname,
      'SAPKundennummer':docItem.SAP_x0020_Kundennummer,
      'DivisionValue':'Automotive',
       'Customer':docItem.Customer0,
      'CustomerClassificationValue':docItem.Customer_x0020_Classification,
      'ResponsibleKAMId':docItem.ResponsibleKAMId,
      'Project':docItem.Project
    };
    const itemUrl = docset.d['__metadata']['uri'];
    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(itemPayload),
      headers: {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "X-RequestDigest": this._webPartContext.pageContext.legacyPageContext.formDigestValue,
        "X-HTTP-Method": "MERGE",
        "If-Match": "*"
      }
    };
    const response2 = await this._webPartContext.spHttpClient.post(itemUrl, SPHttpClient.configurations.v1, spHttpClientOptions);
    console.log(response2)
    if (response2.status === 200) {

      //return true;
    }
   
  }
  public async updateDocumentSetById(docItem:ITodoItem,filteredList:ITodoTaskList,contentTypeId,documentSetname,internalName){
      try {
        const webUrl = this._webPartContext.pageContext.web.absoluteUrl;
      // console.log("updateDocumentSetById,contentTypeId: ",contentTypeId);
      const fileName = this.removeFileExtension(docItem.LinkFilename);
    //  console.log("docItem",docItem);                    ;
      let listItemPayload ;
      if(docItem.Country!=null){
        listItemPayload = {
          //"Title":documentSetname,   
           //"SAP_x0020_Kundennummer0 ": docItem.SAP_x0020_Kundennummer,
          
          "Test_x0020_Division":"Automotive",
          "ContentTypeId":contentTypeId,
          "SAP_x0020_Kundennummer0": docItem.SAP_x0020_Kundennummer,
          "Customer_x0020_Classification":docItem.Customer_x0020_Classification,
          "ResponsibleKAMId":docItem.ResponsibleKAMId,
          "ResponsibleKAMStringId":docItem.ResponsibleKAMStringId,
          "StartDate": docItem.StartDate,
          "Project": docItem.Project,
          "Customer0": docItem.Customer0,
          "Country":{
            "Label":docItem.Country.Label,
            "TermGuid":docItem.Country.TermGuid,
            "WssId":-1
            }
        }
       }else{

          listItemPayload = {
           // "Title": documentSetname,   
           "ContentTypeId":contentTypeId,
            "Test_x0020_Division":"Automotive",
            "SAP_x0020_Kundennummer0": docItem.SAP_x0020_Kundennummer,
            "Customer_x0020_Classification":docItem.Customer_x0020_Classification,
            "ResponsibleKAMId":docItem.ResponsibleKAMId,
            "ResponsibleKAMStringId":docItem.ResponsibleKAMStringId,
            "StartDate": docItem.StartDate,
            "Project": docItem.Project,
            "Customer0": docItem.Customer0
          }
      }
      const spHttpClientOptions: ISPHttpClientOptions = {
        body: JSON.stringify(listItemPayload),
        headers: {
          "Accept": "application/json",
          "Content-Type": "application/json",
          "X-RequestDigest": this._webPartContext.pageContext.legacyPageContext.formDigestValue,
          "X-HTTP-Method": "MERGE",
          "If-Match": "*"
        }
      };
     return  await this._webPartContext.spHttpClient.post(
      `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${this.automotiveListUrl}/${documentSetname}')/ListItemAllFields`,
       SPHttpClient.configurations.v1,
       spHttpClientOptions
    ) 
        
      } catch (error) {
        console.log("Error Occured  :",error);
      }
  }
  public async getCretedDocumentSetIdV2(docItem:ITodoItem,filteredList:ITodoTaskList,contentTypeId,documentSetname,internalName):Promise<any>{
    try {
      const webUrl = this._webPartContext.pageContext.web.absoluteUrl;
      const libraryName =internalName;
      let exceptionLibInternalName = this.getLastPartOfPath(this.automotiveListUrl);
      const libraryUrl = this._webPartContext.pageContext.web.absoluteUrl+"/"+exceptionLibInternalName;
      const folderName = documentSetname;
      const folderContentTypeId = contentTypeId;
      const httpClientOptions: ISPHttpClientOptions = {
          body: JSON.stringify({
            "Title":folderName,
            "Path":libraryUrl,
            "Division": docItem.Division,
            "SAPKundennummer":docItem.SAP_x0020_Kundennummer,
          }),
          headers: {
              "Accept": "application/json;odata=verbose",
              "Slug": `${libraryUrl}/${folderName}|${folderContentTypeId}`,
          }
      };
      return await this._webPartContext.spHttpClient.post(
          `${webUrl}/_vti_bin/listdata.svc/${libraryName}`,
          SPHttpClient.configurations.v1,
          httpClientOptions
      )
          .then((response: SPHttpClientResponse) => {
              if (response.ok) {
                console.log("success fully called _vti_bin/listdata.svc");
                return response.json();
              } else {
                console.log(`Error: ${response.status}`);
              }
              return response.json();
          });
    } catch (error) {
      console.log(error);
          
    }
    return null;
  }
  public escapeParam = (value) => encodeURIComponent(value.replace(/'/g, "''"));

  public async finalUploadDocset(docItem:ITodoItem, libraryName:string,documents: File[], documentSetId: string, contentTypeId:string){
    try {
      const filteredList =SharePointDataProvider._taskLists.filter(item => item.Title ===libraryName)[0];
      libraryName=filteredList.Title;
      let file = documents[documents.length-1]
      const fileBuffer = await this.readFile( documents[documents.length-1]);
      const spHttpClientOptions: ISPHttpClientOptions = {
        body: fileBuffer,
        headers: {
          "Accept": "application/json",
          "Content-Type": "application/json", // Use the MIME type of the file being uploaded
          "X-RequestDigest": this._webPartContext.pageContext.legacyPageContext.formDigestValue,
          "X-HTTP-Method": "POST",
          "X-Microsoft-HTTP-Method": "PUT",
          "If-Match": "*"
        }
      };
    let fileName = file.name;
    let folderName = this.removeFileExtension(file.name);
    const uploadUrl: string = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists/GetById('${filteredList.Id}')/items(${documentSetId})/Folder/files/add(url='${fileName}',overwrite=true)`;
      const response = await this._webPartContext.spHttpClient.post(uploadUrl, SPHttpClient.configurations.v1,spHttpClientOptions)
      .then((response: SPHttpClientResponse) =>{
        if (response.ok){

         return  response.json().then((async x=>{
          try {
           // console.log("finalUploadDocset", x);
            const fileId =parseInt(documentSetId);
            
            //const listItemUrl = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists/GetById('${filteredList.Id}')/items(${fileId})`;
            const listItemUrl = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('/contract_mgmt/cd/Customer%20Documents%20%20Automotive/${folderName}/${fileName}')/ListItemAllFields`;
            
            let listItemPayload ;
            const basePayload = {
              "ContentTypeId": contentTypeId,
              "Title": docItem.Title,
              "Test_x0020_Division": "Automotive",
              "SAP_x0020_Kundennummer0": docItem.SAP_x0020_Kundennummer,
              "OData__Comments": docItem.OData__Comments,
              "Customer0": docItem.Customer0,
              "ResponsibleKAMId": docItem.ResponsibleKAMId,
              "ResponsibleKAMStringId": docItem.ResponsibleKAMStringId,
              "Customer_x0020_Classification": docItem.Customer_x0020_Classification,
              "Project": docItem.Project,
              "Category_x0020_of_x0020_Document0": docItem.Category_x0020_of_x0020_Document,
              "Status_x0020_CD": docItem.Status_x0020_CD,
              "Responible_x0020_Legal" :docItem.Responible_x0020_Legal,
              "ChecklistLink": docItem.ChecklistLink,
              "CurrentStatus": docItem.CurrentStatus,
              "StartDate": docItem.StartDate,
              "Review_x0020_Closed": docItem.Review_x0020_Closed
            };
            
            if (docItem.Country != null && docItem.AMANN_x0020_Company != null) {
              listItemPayload = {
                ...basePayload,
                "Country": {
                  "Label": docItem.Country.Label,
                  "TermGuid": docItem.Country.TermGuid,
                  "WssId": -1
                },
                "AMANN_x0020_Company": {
                  "Label": docItem.AMANN_x0020_Company.Label,
                  "TermGuid": docItem.AMANN_x0020_Company.TermGuid,
                  "WssId": -1
                }
              };
            } else if (docItem.Country != null) {
              listItemPayload = {
                ...basePayload,
                "Country": {
                  "Label": docItem.Country.Label,
                  "TermGuid": docItem.Country.TermGuid,
                  "WssId": -1
                }
              };
            } else if (docItem.AMANN_x0020_Company != null) {
              listItemPayload = {
                ...basePayload,
                "AMANN_x0020_Company": {
                  "Label": docItem.AMANN_x0020_Company.Label,
                  "TermGuid": docItem.AMANN_x0020_Company.TermGuid,
                  "WssId": -1
                }
              };
            } else {
              listItemPayload = basePayload;
            }
            

            //console.log("listItemPayload:",JSON.stringify(listItemPayload));
            const spHttpClientOptionsCall: ISPHttpClientOptions = {
              body: JSON.stringify(listItemPayload),
              headers: {
                "Accept": "application/json",
                "Content-Type": "application/json",
                "X-RequestDigest": this._webPartContext.pageContext.legacyPageContext.formDigestValue,
                "X-HTTP-Method": "MERGE",
                "If-Match": "*"
              }
            };
            return await this._webPartContext.spHttpClient.post(
              listItemUrl,
               SPHttpClient.configurations.v1,
               spHttpClientOptionsCall
            ).then((resultContent=> {
              if(resultContent == undefined || resultContent == null){
                console.log("Issue in finalUploadDocset", resultContent);
              }
              return resultContent;
            }));
            
            // return  await this._webPartContext.spHttpClient.post(listItemUrl, SPHttpClient.configurations.v1, {
            //   headers: {
            //     "Accept": "application/json",
            //     "Content-Type": "application/json",
            //     "X-RequestDigest": this._webPartContext.pageContext.legacyPageContext.formDigestValue,
            //     "X-HTTP-Method": "MERGE",
            //     "If-Match": "*"
            //   },
            //   body: JSON.stringify(listItemPayload)
            // }).then((result => {
            //   return result;
            //   //console.log("listItemResponse ,(result): ",result);
            // }));
            // if (listItemResponse!= undefined && listItemResponse !=null) {
            //   return await listItemResponse;
            // } else {
            //   console.log("listItemResponse== undefined");
            //   console.log(`Error updating list item: ${listItemResponse}`);
            //   return [];
            // }
          } catch (error) {
            console.log("Error occured: Issue in finalUploadDocset",error);
          }
          }))
         // return result;
        }});
      //return response;
    } catch (error) {
      console.log(error);
          return [];
    }
  }
  public async updateUpload(itemId:string,docItem:ITodoItem, libraryName:string, libGuid:string,documents: File[]):Promise<any> {
    try {
      let listUrl= '';
      if(itemId== null || itemId== '' || itemId== undefined){
        console.log(itemId, "item id is missing to update the file props")
        return null;
      }
      let urlreq =  `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists/GetById('${libGuid}')/items(${itemId})`;
      let listItemPayload ;
      let division;
      console.log("docItem",docItem);

      if(libraryName.toLowerCase().includes("consumer")){
        listUrl = this.consumerListUr;
        division = "Consumer";
  
      }else if(libraryName.toLowerCase().includes("apparel")){
        listUrl = this.apparalListUrl;
        division = "Apparel";
      }
      else if(libraryName.toLowerCase().includes("techtex")){
        listUrl = this.techtexListUrl;
        division = "Techtex";
      }
      else if(libraryName.toLowerCase().includes("automotive")){
        listUrl = this.techtexListUrl;
        division = "Automotive";
      }
      if(docItem.Country!=null && docItem.AMANN_x0020_Company!=null){
        listItemPayload = {
        "Test_x0020_Division":division,
        "SAP_x0020_Kundennummer0": docItem.SAP_x0020_Kundennummer,
        "Responsible_x0020_CSCId": docItem.Responsible_x0020_CSC0Id,
        "Responsible_x0020_CSCStringId": docItem.Responsible_x0020_CSC0StringId,
        "Customer0" :docItem.Customer0,
        "Project":docItem.Project,
        "OData__Comments": docItem.OData__Comments,
        "Terms_x0020_of_x0020_Payment":docItem.Terms_x0020_of_x0020_Payment,
        "ResponsibleKAMId":docItem.ResponsibleKAMId,
        "ResponsibleKAMStringId": docItem.ResponsibleKAMStringId,

        "Responible_x0020_Legal": docItem.Responible_x0020_Legal,
        "CurrentStatus" : docItem.CurrentStatus,
        "Status_x0020_CD" :docItem.Status_x0020_CD,
        "StartDate": docItem.StartDate,

        "Country":{
          "Label":docItem.Country.Label,
          "TermGuid":docItem.Country.TermGuid,
          "WssId":-1
          },
          "AMANN_x0020_Company":{
            "Label":docItem.AMANN_x0020_Company.Label,
            "TermGuid":docItem.AMANN_x0020_Company.TermGuid,
            "WssId":-1
        }

    }
      
    }else if(docItem.Country!=null ){
       listItemPayload = {
        // "__metadata": { "type": "SP.Data.YourListNameListItem" },
        "Test_x0020_Division":division,
        "SAP_x0020_Kundennummer0": docItem.SAP_x0020_Kundennummer,
        "Responsible_x0020_CSCId": docItem.Responsible_x0020_CSC0Id,
        "Responsible_x0020_CSCStringId": docItem.Responsible_x0020_CSC0StringId,
        "Customer0" :docItem.Customer0,
        "Project":docItem.Project,
        "OData__Comments": docItem.OData__Comments,
        "Terms_x0020_of_x0020_Payment":docItem.Terms_x0020_of_x0020_Payment,
        "ResponsibleKAMId":docItem.ResponsibleKAMId,
        "ResponsibleKAMStringId": docItem.ResponsibleKAMStringId,

        "Responible_x0020_Legal": docItem.Responible_x0020_Legal,
        "CurrentStatus" : docItem.CurrentStatus,
        "Status_x0020_CD" :docItem.Status_x0020_CD,
        "StartDate": docItem.StartDate,

        "Country":{
          "Label":docItem.Country.Label,
          "TermGuid":docItem.Country.TermGuid,
          "WssId":-1
          }
      }
    }
    else if(docItem.AMANN_x0020_Company!=null ){
       listItemPayload = {
        "Test_x0020_Division":division,
        "SAP_x0020_Kundennummer0": docItem.SAP_x0020_Kundennummer,
        "Responsible_x0020_CSCId": docItem.Responsible_x0020_CSC0Id,
        "Responsible_x0020_CSCStringId": docItem.Responsible_x0020_CSC0StringId,
        "Customer0" :docItem.Customer0,
        "Project":docItem.Project,
        "OData__Comments": docItem.OData__Comments,
        "Terms_x0020_of_x0020_Payment":docItem.Terms_x0020_of_x0020_Payment,
         "ResponsibleKAMId":docItem.ResponsibleKAMId,
         "ResponsibleKAMStringId": docItem.ResponsibleKAMStringId,

         "Responible_x0020_Legal": docItem.Responible_x0020_Legal,
         "CurrentStatus" : docItem.CurrentStatus,
         "Status_x0020_CD" :docItem.Status_x0020_CD,
         "StartDate": docItem.StartDate,

        "AMANN_x0020_Company":{
          "Label":docItem.AMANN_x0020_Company.Label,
          "TermGuid":docItem.AMANN_x0020_Company.TermGuid,
          "WssId":-1
       }
      }
    }
    else{
      listItemPayload = {
        "Test_x0020_Division":division,
        "SAP_x0020_Kundennummer0": docItem.SAP_x0020_Kundennummer,
        "Responsible_x0020_CSCId": docItem.Responsible_x0020_CSC0Id,
        "Responsible_x0020_CSCStringId": docItem.Responsible_x0020_CSC0StringId,
        "Customer0" :docItem.Customer0,
        "Project":docItem.Project,
        "OData__Comments": docItem.OData__Comments,
        "Terms_x0020_of_x0020_Payment":docItem.Terms_x0020_of_x0020_Payment,
         "ResponsibleKAMId":docItem.ResponsibleKAMId,
         "ResponsibleKAMStringId": docItem.ResponsibleKAMStringId,
         "Responible_x0020_Legal": docItem.Responible_x0020_Legal,
         "CurrentStatus" : docItem.CurrentStatus,
         "Status_x0020_CD" :docItem.Status_x0020_CD,
         "StartDate": docItem.StartDate,
      }

    }
    console.log("listItemPayload:",listItemPayload);
      return  await this._webPartContext.spHttpClient.post(urlreq, SPHttpClient.configurations.v1, {
      headers: {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "X-RequestDigest": this._webPartContext.pageContext.legacyPageContext.formDigestValue,
          "X-HTTP-Method": "MERGE",
        "If-Match": "*"
      },
      body: JSON.stringify(listItemPayload)
      }).then((resultContent=> {
        if(resultContent == undefined || resultContent == null){
          console.log("Issue in updateUpload", resultContent);
        }
        return resultContent;
      }));
      
  
      // if(listItemResponse.ok) {
      //   return await listItemResponse.json();
      // }
      // else {
      //   console.log(`Error updating list item: ${listItemResponse.status}`);
      //   return null;;
      // }
      
    } catch (error) {
      console.log(error)
      return null;
    }
   

    

  }
  public async getItemByServerRelativeUrl(serverRelativeUrl):Promise<any> {
    try {
      if(serverRelativeUrl!= undefined && serverRelativeUrl != null){
        const listItemUrl = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelativeUrl}')/ListItemAllFields/ID`;
        return await this._webPartContext.spHttpClient.get(listItemUrl, SPHttpClient.configurations.v1).then((data=>{
          if(data.ok){
            return data.json();
          }else{
            return null;
          }
        }));
      }else{
        return null;
      }
    } catch (error) {
      console.log(error);
      return null;
    }
   

  }
  public async finalUploadV2(docItem:ITodoItem, libraryName:string,libId:string,documents: File[],  contentTypeId:string):Promise<any> {
    let listUrl= null;
    let division;
    if(libraryName.toLowerCase().includes("consumer")){
      listUrl = this.consumerListUr;
      division = "Consumer";

    }else if(libraryName.toLowerCase().includes("apparel")){
      listUrl = this.apparalListUrl;
      division = "Apparel";
    }
    else if(libraryName.toLowerCase().includes("techtex")){
      listUrl = this.techtexListUrl;
      division = "Techtex";
    }
    else if(libraryName.toLowerCase().includes("automotive")){
      listUrl = this.techtexListUrl;
      division = "Automotive";
    }
    const file = documents[documents.length-1]
    const fileBuffer = await this.readFile(file);

    // Define the file name and content
    const fileName = file.name;


    // Get the current context's web URL
    const webUrl = this._webPartContext.pageContext.web.absoluteUrl;
    // const fileMetadata = {
    //   Title: fileName,
    //   ServerRelativeUrl: `${webUrl}${listUrl}/${fileName}`,
    //   ContentTypeId:  contentTypeId

    // };

    var targetUrl = this._webPartContext.pageContext.web.absoluteUrl + "/" + libraryName;  

    


   // var url = `${webUrl}/_api/web/lists/getByTitle('${libraryName}')/RootFolder/Files/Add(url='${fileName}', overwrite=true)`;
    var url = `${webUrl}/_api/web/lists/GetById('${libId}')/RootFolder/Files/Add(url='${fileName}', overwrite=true)`;


    // Prepare the request headers
    const spHttpClientOptions: ISPHttpClientOptions = {
      body: fileBuffer,
      headers: {
        "Accept": "application/json",
        "Content-Type": "application/json", // Use the MIME type of the file being uploaded
        "X-RequestDigest": this._webPartContext.pageContext.legacyPageContext.formDigestValue,
        "X-HTTP-Method": "POST",
        "X-Microsoft-HTTP-Method": "PUT",
        "If-Match": "*"
      }
    };
    // Make the POST request to upload the file
      return await this._webPartContext.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions).then((response: SPHttpClientResponse) => {
      if (response.ok) {
        console.log("File added successfully.:finalv2",response);
        return response.json();
      } else {
        //return response.json();
        console.log("Error uploading the file:", response);
      }
    }).catch((error) => {
      console.log(error)
      console.log("Error:", error);
    });
  }
  public getLists = async (): Promise<ITodoTaskList[]> => {
    return await this._getLists();
  };
  public async getPermissions():Promise<boolean>{
    // var isV = await this.isVisitor(this._webPartContext);
    // var isW = await this.isOwner(this._webPartContext);
   
    // console.log("IV",isV);
    // console.log("IW",isW);

     return await this.isUserSiteMemberWithEditPermissions(this._webPartContext);
   //  return x;
  }
  public  isUserSiteMemberWithEditPermissions = async (context) => {
    const siteUrl = context.pageContext.web.absoluteUrl;
    try {

      const response: SPHttpClientResponse = await context.spHttpClient.get(
         `${siteUrl}/_api/web/CurrentUser?$expand=Groups&$select=Id,Groups/Id,Groups/Title,Groups/CanEdit`,SPHttpClient.configurations.v1);
      //  `${siteUrl}/_api/web/RoleAssignments?$expand=Member&$filter=Member/LoginName eq '${currentUserLoginName}' and (RoleDefinitionBindings/Name eq 'Contribute' or RoleDefinitionBindings/Name eq 'Edit' )  `, SPHttpClient.configurations.v1);
      if (response.ok) {
        const user = await response.json();
        const groups = user.Groups;
        console.log("Current user groups ",groups);
        const filteredList =groups.filter(item =>  item.Title.toLowerCase().includes("member") ||   item.Title.toLowerCase().includes("owner") || item.Title.toLowerCase().includes("visitor"))[0];
        console.log("after filter in ismem", filteredList);
        if(filteredList!=undefined){
          return true;
        }

        return false;
      } else {
        console.log(`Error: ${response.status}`);
        return false;
      }
    } catch (error) {
      console.log('Error:', error);
      return false;
    }
  };
  public  isVisitor = async (context) => {
    const siteUrl = context.pageContext.web.absoluteUrl;
    var flag= false;
    try {

      const response: SPHttpClientResponse = await context.spHttpClient.get(
         `${siteUrl}/_api/web/CurrentUser?$expand=Groups&$select=Id,Groups/Id,Groups/Title,Groups/CanEdit`,SPHttpClient.configurations.v1);
      //  `${siteUrl}/_api/web/RoleAssignments?$expand=Member&$filter=Member/LoginName eq '${currentUserLoginName}' and (RoleDefinitionBindings/Name eq 'Contribute' or RoleDefinitionBindings/Name eq 'Edit' )  `, SPHttpClient.configurations.v1);
      if (response.ok) {
        const user = await response.json();
        const groups = user.Groups;
        console.log("Current user groups ",groups);
        const filteredList =groups.filter(item =>  item.Title.toLowerCase().includes("customer documents visitors"))[0];
        console.log("after filter in isv" ,filteredList);
        if(filteredList!=undefined){        
          console.log("Isvisitor: " ,filteredList);
          flag= true;
          return flag;
        }else{
          console.log("Is Visitor: no data filtered " ,);
          return flag;
        }
      } else {
        console.log(`Error: ${response.status}`);
        return flag;
      }
    } catch (error) {
      console.log('Error:', error);
      return flag;
    }
  };
  public async isOwner(context)
  {
    var flag = false;
    const siteUrl = context.pageContext.web.absoluteUrl;
    try {

      const response: SPHttpClientResponse = await context.spHttpClient.get(
         `${siteUrl}/_api/web/CurrentUser?$expand=Groups&$select=Id,Groups/Id,Groups/Title,Groups/CanEdit`,SPHttpClient.configurations.v1);
      //  `${siteUrl}/_api/web/RoleAssignments?$expand=Member&$filter=Member/LoginName eq '${currentUserLoginName}' and (RoleDefinitionBindings/Name eq 'Contribute' or RoleDefinitionBindings/Name eq 'Edit' )  `, SPHttpClient.configurations.v1);
      if (response.ok) {
        const user = await response.json();
        const groups = user.Groups;
        console.log("Current user groups ",groups);
        const filteredList =groups.filter(item =>  item.Title.toLowerCase().includes("contract management owners"))[0];
        console.log("after filter in isw" ,filteredList);
       //console.log("isOwner: " ,filteredList);
        if(filteredList!=undefined){
          console.log("IsOwner: " ,filteredList);
          flag= true;
          return flag;
        }else{
          console.log("IsOwner: no data filtered " ,);
          return flag;
        }
        
      } else {
        console.log(`Error: ${response.status}`);
        return flag;
      }
    } catch (error) {
      console.log('Error:', error);
      return flag;
    }
  };
  private async checkGroupMembership(groupName: string): Promise<boolean> {
    try {
      const group = await sp.web.siteGroups.getByName(groupName).users
        .filter(`Id eq ${this._webPartContext.pageContext.legacyPageContext.userId}`)
        .get();
        this.printStatus("CheckcheckGroupMembership, groupName: "+groupName +"destLibIds :", group);
      return group.length > 0;
    } catch (error) {
      console.error(`Error checking group membership: ${error}`);
      return false;
    }
  }
  public  printStatus (msg:string, obj:any){
    console.log(msg,obj);
  }
  private async _getLists(): Promise<ITodoTaskList[]> {
    let filter = "Hidden eq false and ";
    filter += "BaseType eq 1 and BaseTemplate eq 101";
    const endpointUrl = `${this.getSiteUrl()}/_api/web/lists?$filter=${filter}`;
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      }
    };
    const response: SPHttpClientResponse = await this._webPartContext.spHttpClient.get(endpointUrl, SPHttpClient.configurations.v1,options);
    return await response.json().then((json: { value: ITodoTaskList[] }) => {
      return  SharePointDataProvider._taskLists = json.value.map( (task: ITodoTaskList) => {
       let test:ITodoTaskList =task;
        return test;
      });
    });
  }
  public getDocs(): Promise<ITodoItem[]> {
    return this._getDocs(this.webPartContext.spHttpClient);
  }
  public getDoc(itemId:string): Promise<ITodoItem> {
    return this._getDoc(itemId);
  }
  private async _getDocs(requester: SPHttpClient): Promise<ITodoItem[]> {
   var res = await this.isOwner(this._webPartContext);
   console.log("_getDocs is owner"+ this._webPartContext.pageContext.user.displayName, res);

    var sortOrder = "desc";
    var queryString: string =`?$select=Title,HasUniqueRoleAssignments,Terms_x0020_of_x0020_Payment0,Status_x0020_CD,StartDate,Review_x0020_Closed,Responible_x0020_Legal,ChecklistLink,CurrentStatus,Project,SAP_x0020_Kundennummer,OData__Comments,Responsible_x0020_CSC0Id,Responsible_x0020_CSC0StringId,Category_x0020_of_x0020_Document,Customer_x0020_Classification,Responsible_x0020_CSCId,Responsible_x0020_CSCStringId,ResponsibleKAMStringId,ResponsibleKAMId,Customer0,Turnover_x0020__x0028_Prev_x002e_Year_x0029_,Responible_x0020_Legal,Country,AMANN_x0020_Company,Priority,Deadline_x0020_Customer,Received0,Incoterms_x0020__x0028_current_x0029_,Label0,Trunover_x0020__x0028_Prev_x002e_Year_x002d_YTD_x0029_,Customer_x0020_Classification,ID,Created,UniqueId,FileRef,FileDirRef,LinkFilename,LinkFilename2,ServerUrl,FileLeafRef,ContentTypeId,Author/Title,Editor/Title&$expand=Editor/Title&expand=Modified/ID,Modified/Title,File/LinkingUrl&$expand=File&$expand=Author/Title&$top=3000&$orderby=Created%20desc`;
   
    if(!res){
       queryString =`?$select=Title,HasUniqueRoleAssignments,Terms_x0020_of_x0020_Payment0,Status_x0020_CD,StartDate,Review_x0020_Closed,Responible_x0020_Legal,ChecklistLink,CurrentStatus,Project,SAP_x0020_Kundennummer,OData__Comments,Responsible_x0020_CSC0Id,Responsible_x0020_CSC0StringId,Category_x0020_of_x0020_Document,Customer_x0020_Classification,Responsible_x0020_CSCId,Responsible_x0020_CSCStringId,ResponsibleKAMStringId,ResponsibleKAMId,Customer0,Turnover_x0020__x0028_Prev_x002e_Year_x0029_,Responible_x0020_Legal,Country,AMANN_x0020_Company,Priority,Deadline_x0020_Customer,Received0,Incoterms_x0020__x0028_current_x0029_,Label0,Trunover_x0020__x0028_Prev_x002e_Year_x002d_YTD_x0029_,Customer_x0020_Classification,ID,Created,UniqueId,FileRef,FileDirRef,LinkFilename,LinkFilename2,ServerUrl,FileLeafRef,ContentTypeId,Author/Title,Editor/Title&$expand=Editor/Title&expand=Modified/ID,Modified/Title,File/LinkingUrl&$expand=File&$expand=Author/Title&$filter=Author/Title eq '${this._webPartContext.pageContext.user.displayName}'&$top=3000&$orderby=Created%20desc`;
    }
    //const queryString: string =`?$select=Title,ID,Created,UniqueId,FileRef,FileDirRef,LinkFilename,LinkFilename2,ServerUrl,FileLeafRef,ContentTypeId,Author/Title,Editor/Title&$expand=Editor/Title&expand=Modified/ID,Modified/Title&$expand=Author/Title,File/LinkingUrl&$expand=File&$top=3000&$orderby=Created `+sortOrder;
   // const queryString: string =`?$select=Title,Responsible_x0020_CSC,ResponsibleKAM,Country,AMANN_x0020_Company,Priority,time_x0020_customer,Received,Incoterms_x0020__x0028_currently_x0029_,Terms_x0020_of_x0020_payment,Label0,Trunover_x0020__x0028_Prev_x002e_Year_x002d_YTD_x0029_,Comments,Customer_x0020_Classification,ID,Created,UniqueId,FileRef,FileDirRef,LinkFilename,LinkFilename2,ServerUrl,FileLeafRef,ContentTypeId,Author/Title,Editor/Title&$expand=Editor/Title&expand=Modified/ID,Modified/Title&$expand=Author/Title,File/LinkingUrl&$expand=File&$top=3000&$orderby=Created `+sortOrder;
   
    const queryUrl: string = this._listItemsUrl + queryString;
    // const requestOptions: ISPHttpClientOptions = {
    //   headers: {
    //     'Accept': 'application/octet-stream'
    //   }
    // };
    return await requester.get(queryUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((json: { value: ITodoItem[] }) => {
        return json.value.map((task: ITodoItem) => {
          task.DefaultEditUrl = `${this._webPartContext.pageContext.web.absoluteUrl}/_layouts/15/listform.aspx?PageType=6&ListId=${this.selectedList.Id}&ID=${task.Id}&RootFolder=*`;
          if(res!= undefined && res != null && res==true){
            task.IsMovePermission = true;
          }else{
            task.IsMovePermission= false;
          }
          return task;
        });
      });
  }
  private _getDoc(itemId: string): Promise<ITodoItem> {
    const queryString: string =`(${itemId})?$select=Title,Terms_x0020_of_x0020_Payment0,Status_x0020_CD,StartDate,Review_x0020_Closed,Responible_x0020_Legal,ChecklistLink,CurrentStatus,Project,SAP_x0020_Kundennummer,OData__Comments,Responsible_x0020_CSC0Id,Responsible_x0020_CSC0StringId,Category_x0020_of_x0020_Document,Customer_x0020_Classification,Responsible_x0020_CSCId,Responsible_x0020_CSCStringId,ResponsibleKAMStringId,ResponsibleKAMId,Customer0,Turnover_x0020__x0028_Prev_x002e_Year_x0029_,Responible_x0020_Legal,Country,AMANN_x0020_Company,Priority,Deadline_x0020_Customer,Received0,Incoterms_x0020__x0028_current_x0029_,Label0,Trunover_x0020__x0028_Prev_x002e_Year_x002d_YTD_x0029_,Customer_x0020_Classification,ID,Created,UniqueId,FileRef,FileDirRef,LinkFilename,LinkFilename2,ServerUrl,FileLeafRef,ContentTypeId,Author/Title,Editor/Title&$expand=Editor/Title&expand=Modified/ID,Modified/Title&$expand=Author/Title,File/LinkingUrl&$expand=File`;
   console.log("Calling culprit: :", itemId);
    //const queryString: string =`?$select=Title,Project,SAP_x0020_Kundennummer,Category_x0020_of_x0020_Document,OData__Comments,Responsible_x0020_CSC0Id,Responsible_x0020_CSC0StringId,Customer_x0020_Classification,Responsible_x0020_CSCId,Responsible_x0020_CSCStringId,ResponsibleKAMStringId,ResponsibleKAMId,Customer0,Turnover_x0020__x0028_Prev_x002e_Year_x0029_,Responible_x0020_Legal,Country,AMANN_x0020_Company,Priority,Deadline_x0020_Customer,Received0,Incoterms_x0020__x0028_current_x0029_,Label0,Trunover_x0020__x0028_Prev_x002e_Year_x002d_YTD_x0029_,Customer_x0020_Classification,ID,Created,UniqueId,FileRef,FileDirRef,LinkFilename,LinkFilename2,ServerUrl,FileLeafRef,ContentTypeId,Author/Title,Editor/Title&$expand=Editor/Title&expand=Modified/ID,Modified/Title&$expand=Author/Title,File/LinkingUrl&$expand=File&$filter=Id eq '${itemId}'`;
    //const queryString: string =`?$select=*&$top=3000&$orderby=Created `+sortOrder;
    const queryUrl: string = this._listItemsUrl + queryString;
    return this._webPartContext.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        console.log("culprit : ",response);
        return response.json()
        .then((dataitm =>{
          console.log("culprit : ",dataitm)
          let itm:ITodoItem = dataitm;
          return itm;
        }))
      });
  }
  public createDoc(documents: File[]): Promise<ITodoItem[]> {
    return this
      ._createDoc(documents, this.webPartContext.spHttpClient)
      .then(_resp => {
        return this.getDocs();
      });
  }
  private _createDoc(documents: File[], client: SPHttpClient): Promise<SPHttpClientResponse> {
    let len = documents.length-1;
    let file:File = documents[len];
    const fileBuffer = new Blob([file], { type: file.type });

    const spOpts: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json',
      },
      body: fileBuffer,
    };
    let url = this._docItemsUrl + `/RootFolder/Files/Add(url='${file.name}', overwrite=true)?$expand=ListItemAllFields`;
        return client.post(
          url,
          SPHttpClient.configurations.v1,
          spOpts
        );
  }



  private _removePermissions(documentName: string, client: SPHttpClient): Promise<void> {
    // Form the URL to break role inheritance for the specific document
    const breakRoleInheritanceUrl = `${this._docItemsUrl}/RootFolder/Files('${documentName}')/ListItemAllFields/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`;
  
    const spOpts: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json',
      },
    };
  
    // Break role inheritance for the specific document
    return client.post(breakRoleInheritanceUrl, SPHttpClient.configurations.v1, spOpts)
      .then(() => {
        console.log(`Removed unique permissions for ${documentName}`);
      });
  }
  
  private _inheritPermissions(documentName: string, client: SPHttpClient): Promise<void> {
    const editPermissionRoleDefId = 1073741827;
    const siteOwnersPrincipalId = 4;
    // Form the URL to assign permissions to the "Site Owners" group for the specific document
    const assignPermissionsUrl = `${this._docItemsUrl}/RootFolder/Files('${documentName}')/ListItemAllFields/roleassignments/addroleassignment(principalid=${siteOwnersPrincipalId}, roledefid=${editPermissionRoleDefId})`;
  
    const spOpts: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json',
      },
    };
  
    // Inherit permissions for the specific document from the library
    return client.post(assignPermissionsUrl, SPHttpClient.configurations.v1, spOpts)
      .then(() => {
        console.log(`Inherited permissions for ${documentName} from the library`);
      });
  }
  
 
  public deleteDoc(itemDeleted: ITodoItem): Promise<ITodoItem[]> {
    return this
      ._deleteDoc(itemDeleted, this.webPartContext.spHttpClient)
      .then(_ => {
        return this.getDocs();
      });
  }
  private _deleteDoc(item: ITodoItem, client: SPHttpClient): Promise<SPHttpClientResponse> {
    const itemDeletedUrl: string = `${this._listItemsUrl}(${item.Id})`;

    const headers: Headers = new Headers();
    headers.append('If-Match', '*');

    return client.fetch(itemDeletedUrl,
      SPHttpClient.configurations.v1,
      {
        headers,
        method: 'DELETE'
      }
    );
  }
  public getFileTypeByExtension(extension) {
    // Remove the dot (.) if present in the extension
    extension = extension.replace(".", "");

    // Define mappings of file extensions to MIME types
    var extensionToMimeTypeMap = {
      txt: "text/plain",
      html: "text/html",
      css: "text/css",
      js: "application/javascript",
      json: "application/json",
      xml: "application/xml",
      jpg: "image/jpeg",
      jpeg: "image/jpeg",
      png: "image/png",
      pdf: "application/pdf",
      docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
      // Add more mappings as needed
    };

    // Look up the MIME type based on the extension
    var mimeType = extensionToMimeTypeMap[extension.toLowerCase()];

    // Return the MIME type or a default value if not found
    return mimeType || "application/octet-stream";
  }
  public get selectedList(): ITodoTaskList {
    return this._selectedList;
  }
  public set webPartContext(value: IWebPartContext) {
    this._webPartContext = value;
    this._listsUrl = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists`;
  }
  public get webPartContext(): IWebPartContext {
    return this._webPartContext;
  }
  public getTaskLists(): Promise<ITodoTaskList[]> {
    const queryUrl: string = this._listsUrl + '?$select=Title,Id,ListItemEntityTypeFullName';

    return this._webPartContext.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((json: { value: ITodoTaskList[] }) => {
        return this._taskLists = json.value;
      });
  }
  public getItems(): Promise<ITodoItem[]> {

    return this._getItems(this.webPartContext.spHttpClient);

  }
  public createItem(title: string): Promise<ITodoItem[]> {
    return this
      ._createItem(title, this.webPartContext.spHttpClient)
      .then(_ => {
        return this.getItems();
      });
  }
  public deleteItem(itemDeleted: ITodoItem): Promise<ITodoItem[]> {
    return this
      ._deleteItem(itemDeleted, this.webPartContext.spHttpClient)
      .then(_ => {
        return this.getItems();
      });
  }
  public updateItem(itemUpdated: ITodoItem): Promise<ITodoItem[]> {
    return this
      ._updateItem(itemUpdated, this.webPartContext.spHttpClient)
      .then(_ => {
        return this.getItems();
      });
  }
  private _getItems(requester: SPHttpClient): Promise<ITodoItem[]> {
    const queryString: string = `?$select=Id,Title,PercentComplete`;
    const queryUrl: string = this._listItemsUrl + queryString;

    return requester.get(queryUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((json: { value: ITodoItem[] }) => {
        return json.value.map((task: ITodoItem) => {
          return task;
        });
      });
  }
  private _createItem(title: string, client: SPHttpClient): Promise<SPHttpClientResponse> {
    const body: {} = {
      '@data.type': `${this._selectedList.EntityTypeName}`,
      'Title': title
    };

    return client.post(
      this._listItemsUrl,
      SPHttpClient.configurations.v1,
      { body: JSON.stringify(body) }
    );
  }
  
  private _deleteItem(item: ITodoItem, client: SPHttpClient): Promise<SPHttpClientResponse> {
    const itemDeletedUrl: string = `${this._listItemsUrl}(${item.Id})`;

    const headers: Headers = new Headers();
    headers.append('If-Match', '*');

    return client.fetch(itemDeletedUrl,
      SPHttpClient.configurations.v1,
      {
        headers,
        method: 'DELETE'
      }
    );
  }
  private _updateItem(item: ITodoItem, client: SPHttpClient): Promise<SPHttpClientResponse> {

    const itemUpdatedUrl: string = `${this._listItemsUrl}(${item.Id})`;

    const headers: Headers = new Headers();
    headers.append('If-Match', '*');

    const body: {} = {
      '@data.type': `${this._selectedList.EntityTypeName}`,
      //'PercentComplete': item.PercentComplete
    };

    return client.fetch(itemUpdatedUrl,
      SPHttpClient.configurations.v1,
      {
        body: JSON.stringify(body),
        headers,
        method: 'PATCH'
      }
    );
  }
  public static convertEscapedString(title: string): string {
    return title.replace(/\s/g, '');
  }
  public removeFileExtension(fileName) {
    const lastDotIndex = fileName.lastIndexOf('.');
    return fileName.substring(0, lastDotIndex);
  }
  public static getLastSegmentFromUrl(url) {
    const segments = url.split("/");
    return (segments[segments.length - 3])// Get the second-to-last segment
  }
  public  minimizeText(text, maxCharacters) {
    if( text!=undefined){
      if (text.length <= maxCharacters) {
        return text;
      }
      const minimizedText = text.slice(0, maxCharacters)+'...';
      return minimizedText;
    }
    return '';
  }
  public getFileExtension(filename) {
    if(filename!= undefined){
      var parts = filename.split('.');
      return parts[parts.length - 1];
    }
    return '';
    
  }
  public removeDotswithUnderScrol(name) {
    if(name!= undefined){
      var modifiedName = name.replace(/[^\w\s]/g, '_');
      return modifiedName;
      //console.log(modifiedName);  // Output: john_doe
    }else{
      return '';
    }
   
  }
  private async readFile(file: File): Promise<ArrayBuffer> {
    return new Promise<ArrayBuffer>((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = (event: any) => {
        resolve(event.target.result);
      };
      reader.onerror = (event: any) => {
        reject(event.target.error);
      };
      reader.readAsArrayBuffer(file);
    });
  }
  showAlert(type,msg,defaultmsg){
    Swal(type,msg, defaultmsg);
  };
  public getSiteUrl(): string {
    return this._webPartContext.pageContext.web.absoluteUrl;
  }
  public  convertText(text) {
    // Remove spaces and special characters using regular expressions
    let cleanedText = text.replace(/[^a-zA-Z0-9]/g, '');
    
    // Split the cleaned text into individual words
    let words = cleanedText.split('_x0020_');
    
    // Remove any empty words
    words = words.filter(word => word.length > 0);
    
    // Join the words together without any separators
    let convertedText = words.join('');
    
    return convertedText;
  }
  public  getLastPartOfPath(path) {
    // Split the path by "/"
    var parts = path.split("/");
  
    // Remove any empty parts
    parts = parts.filter(function(part) {
      return part !== "";
    });
  
    // Retrieve the last part of the path
    var lastPart = parts.pop();
  
    return lastPart;
  }
}
