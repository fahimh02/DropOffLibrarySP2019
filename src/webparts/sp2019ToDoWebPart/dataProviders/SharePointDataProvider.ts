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

  }

  public async uploadItems(item:ITodoItem, selectedLibrary:string, createDocumentSet:boolean): Promise<ITodoItem>{
    if(selectedLibrary.toLowerCase().includes("automotive") || selectedLibrary.toLowerCase().includes("b5f11715-9d34-416c-9d0c-bd055ee95400")){
      return await this._uploadItemsInDocSet(item,selectedLibrary,createDocumentSet)
      .then((itemresponse =>{
        return itemresponse;
        }
      ))
    }else{
      return await this._uploadItemInDoclib(item,selectedLibrary)
    .then((itemresponse =>{
      return itemresponse;
      }
    ))
    }
  }
  public async _uploadItemsInDocSet(item:ITodoItem, selectedLibrary:string, createDocumentSet:boolean):Promise<ITodoItem> {
    let statusReq:boolean = false;
    try {
          const documentContent:any  = await this.getFileFormServer(item);
         if(documentContent !== undefined)
         {
           const filteredList =SharePointDataProvider._taskLists.filter(item => item.Id ===selectedLibrary)[0];
           let selectedLibraryTitle = filteredList.Title;
           this.convertText( filteredList.Title);

           var internalName= SharePointDataProvider.convertEscapedString(filteredList.EntityTypeName);
           if(internalName =="Prototyp_x0020_Automotive_x0020_Improved"){
           internalName = 'PrototypAutomotiveImproved';
           }
         
           if(filteredList.EntityTypeName == "Customer_x0020_Documents_x0020__x0020_Automotive"){
            internalName = "CustomerDocumentsAutomotive";
           }

           console.log(filteredList, this._listsUrl);
           var cttype  = await this.getAvaiableContentTypesByDocRef(selectedLibraryTitle);
           
           var docsetContentTypeId='' ;
           var documentContentTypeId='' ;
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
             //let test = await this.moveFileByPath("","",item.LinkFilename,true,true);
             var docset = await this.getCretedDocumentSetIdV2(item,filteredList,docsetContentTypeId,documentSetname,internalName);
             if(docset!=undefined && docset!=null && docset.d !=undefined && docset.d.Id!=undefined ){
          
              let docsetId = docset.d.Id;
          

              var updateDocSet = await this.updateDocumentSetById(docset,item,filteredList,docsetContentTypeId,documentSetname,internalName);
        
              //this.moveFile(item.ServerUrl,"");
             
             let docfile = documentContent as ArrayBuffer;
             let file:File = documentContent;
             const selectedFiles:File[]= [];
             selectedFiles.push(file);
       
             let ext = this.getFileExtension(item.LinkFilename);
             let fileName = documentSetname;
         
             var modifiedStr = fileName.replace(/\s+\./g, '.');// additional space
             var modifiedStr = modifiedStr.replace(/-/g, '');// remove hyfen

             fileName = this.minimizeText(modifiedStr,30)+'.'+ext;
             
            //  let destFileUrl = filteredList.ParentWebUrl+"/"+filteredList.EntityTypeName+"/"+documentSetname+"/"+fileName;
            //  console.log(destFileUrl);

           //  let responseupload = await this.moveFile(item.ServerUrl,destFileUrl,documentContentTypeId);
             
             //console.log(responseupload);
            let responseupload = await this.finalUploadDocset(item,selectedLibraryTitle,selectedFiles,docsetId,documentContentTypeId);
           // let responseupload = undefined;

            if(responseupload !=undefined){
              console.log("responseupload",responseupload);
              this.showAlert("Success!", file.name+" moved sucessfully!","success");
              statusReq = true;
              return item;
            }else if(docset.t.nativeResponse!= undefined){
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
    
               this.showAlert("Error!", "There is a problem creating the document set! Name might exists.code :"+docset.t.nativeResponse.status,"error");
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
           //console.log(documentContent);
           const filteredList =SharePointDataProvider._taskLists.filter(item => item.Id ===selectedLibrary)[0];
   
           selectedLibrary = filteredList.Title;
           this.convertText( filteredList.Title);

           var internalName= SharePointDataProvider.convertEscapedString(filteredList.EntityTypeName);
           if(internalName =="Prototyp_x0020_Automotive_x0020_Improved"){
           internalName = 'PrototypAutomotiveImproved';
           }
           if(filteredList.Id =="6ac7c5a8-9cff-4d31-bdac-8186a2d198ab"){
           internalName = 'Customer%20Documents%20Apparel';
           //documentSetContentTypeName = "CD Doc Set Auto";
           }

           var cttype  = await this.getAvaiableContentTypesByDocRef(selectedLibrary);
           
          
           var documentContentTypeId='' ;
           if(cttype!= undefined &&  cttype.value.length>0)
           {
             let contentTypes = cttype.value;
             this.documentContentTypeName = "customer document";
         
             const documentContentType = contentTypes.filter(item => item.Name.toLowerCase().includes(this.documentContentTypeName))[0];

           
             documentContentTypeId = documentContentType["Id"]["StringValue"];

             let documentSetname = this.removeFileExtension(item.LinkFilename);
             //let test = await this.moveFileByPath("","",item.LinkFilename,true,true);
            
            
             let docfile = documentContent as ArrayBuffer;
             let file:File = documentContent;
             const selectedFiles:File[]= [];
             selectedFiles.push(file);
             let ext = this.getFileExtension(item.LinkFilename);
             let fileName = documentSetname;
             var modifiedStr = fileName.replace(/\s+\./g, '.');// additional space
             var modifiedStr = modifiedStr.replace(/-/g, '');// remove hyfen

             fileName = this.minimizeText(modifiedStr,30)+'.'+ext;
             
            // let destFileUrl = filteredList.ParentWebUrl+"/"+internalName+"/"+fileName;


             //let responseupload = await this.moveFile(item.ServerUrl,destFileUrl,documentContentTypeId);
             
             //console.log(responseupload);


            let responseupload = await this.finalUploadV2(item,filteredList.Title,selectedFiles,documentContentTypeId);

            console.log("responseupload", responseupload)
            if(responseupload !=undefined){
              let uploadedItem =  await this.getItemByServerRelativeUrl(responseupload.ServerRelativeUrl);
              console.log("getItemByServerRelativeUrl", uploadedItem)
              if(uploadedItem!= undefined && uploadedItem.value!= undefined && uploadedItem!= null){
                let id = uploadedItem.value;
                let updateupload = await this.updateUpload(id,item,filteredList.Title,selectedFiles);
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
      console.log("world finish from upload");
      
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

  public async getAvaiableContentTypesByDocRef(documentLibraryName){
    try {
      let reqURL = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${documentLibraryName}')/contenttypes?$select=Id,Name`;
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
  private async getRequestDigestValue(): Promise<string> {
    const webUrl = this._webPartContext.pageContext.web.absoluteUrl;
    const endpointUrl = `${webUrl}/_api/contextinfo`;

    const requestOptions: any = {
      method: 'POST',
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json',
      },
    };

    try {
      const response = await this._webPartContext.spHttpClient.fetch(endpointUrl, SPHttpClient.configurations.v1, requestOptions);
      const data = await response.json();

      if (response.ok && data && data.FormDigestValue) {
        return data.FormDigestValue;
      } else {
        throw new Error('Failed to retrieve request digest.');
      }
    } catch (error) {
      throw new Error('Error retrieving request digest: ' + error);
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
  public async updateDocumentSetById(docset,docItem:ITodoItem,filteredList:ITodoTaskList,contentTypeId,documentSetname,internalName){

      let docsetId = docset.d.Id;
      const webUrl = this._webPartContext.pageContext.web.absoluteUrl;
      const libraryName =filteredList.Title   ; 
    //  console.log("docItem",docItem);                    ;
      let listItemPayload ;
      if(docItem.Country!=null){
        listItemPayload = {
          //"Title":documentSetname,   
           //"SAP_x0020_Kundennummer0 ": docItem.SAP_x0020_Kundennummer,
          
          "Test_x0020_Division":"Automotive",
          "SAP_x0020_Kundennummer0": docItem.SAP_x0020_Kundennummer,
          "Customer_x0020_Classification":docItem.Customer_x0020_Classification,
          "ResponsibleKAMId":docItem.ResponsibleKAMId,
          "ResponsibleKAMStringId":docItem.ResponsibleKAMStringId,
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
            "Test_x0020_Division":"Automotive",
            "SAP_x0020_Kundennummer0": docItem.SAP_x0020_Kundennummer,
            "Customer_x0020_Classification":docItem.Customer_x0020_Classification,
            "ResponsibleKAMId":docItem.ResponsibleKAMId,
            "ResponsibleKAMStringId":docItem.ResponsibleKAMStringId,
            "Project": docItem.Project,
            "Customer0": docItem.Customer0
          }

      }

    const digestValue = await this.getRequestDigestValue();

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
  }

  public async getCretedDocumentSetIdV2(docItem:ITodoItem,filteredList:ITodoTaskList,contentTypeId,documentSetname,internalName):Promise<any>{

    try {
      const webUrl = this._webPartContext.pageContext.web.absoluteUrl;
      const libraryName =internalName;
      this.automotiveListUrl
      let exceptionLibInternalName = this.getLastPartOfPath(this.automotiveListUrl);

      // const folderPayload = { 
      //   ContentTypeId: contentTypeId,
      //   SAP_x0020_Kundennummer: "test_SAP_x0020_Kundennummer",
  
      //   Customer_x0020_Group_x0020_Company: "TEST"

      //   /* your folder payload */ 
      // };
      const libraryUrl = this._webPartContext.pageContext.web.absoluteUrl+"/"+exceptionLibInternalName;
      const folderName = documentSetname;
      const folderContentTypeId = contentTypeId;
      
      const httpClientOptions: ISPHttpClientOptions = {
          body: JSON.stringify({
            "Title":folderName,
            "Path":libraryUrl,
            "Division": docItem.Division,
            "SAPKundennummer":docItem.SAP_x0020_Kundennummer
          }),
          headers: {
              "Accept": "application/json;odata=verbose",
              "Slug": `${libraryUrl}/${folderName}|${folderContentTypeId}`,
          }
      };
      
      return await this._webPartContext.spHttpClient.post(
            `${webUrl}/_vti_bin/listdata.svc/${libraryName}`,
          // `${webUrl}/_api/web/lists/getbytitle('${filteredList.Title}')/rootfolder/folders`,
          
          // `${webUrl}/_api/web/lists/getbytitle('${libraryName}')/items(${data.d.Id})`,
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
  public async finalUploadDocset(docItem:ITodoItem, libraryName:string,documents: File[], documentSetId: string, contentTypeId:string) {
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
    const uploadUrl: string = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${libraryName}')/items(${documentSetId})/Folder/files/add(url='${fileName}',overwrite=true)`;
      const response = await this._webPartContext.spHttpClient.post(uploadUrl, SPHttpClient.configurations.v1,spHttpClientOptions)
      .then((response: SPHttpClientResponse) =>{
        if (response.ok){
         return  response.json().then((async x=>{
            const fileId =parseInt(documentSetId) +1;
            const listItemUrl = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${libraryName}')/items(${fileId})`;
            let listItemPayload ;
            if(docItem.Country!=null && docItem.AMANN_x0020_Company!=null){
              listItemPayload = {
            // "__metadata": { "type": "SP.Data.YourListNameListItem" },
            "ContentTypeId": contentTypeId,
            "Title": docItem.Title,
            "Test_x0020_Division":"Automotive",
            "SAP_x0020_Kundennummer0": docItem.SAP_x0020_Kundennummer,
            "OData__Comments":docItem.OData__Comments,
            "Customer0":docItem.Customer0,
            "ResponsibleKAMId":docItem.ResponsibleKAMId,
            "ResponsibleKAMStringId": docItem.ResponsibleKAMStringId,
            "Customer_x0020_Classification":docItem.Customer_x0020_Classification,
            "Project":docItem.Project,
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
              "ContentTypeId": contentTypeId,
              "Title": docItem.Title,
              "Test_x0020_Division":"Automotive",
              "SAP_x0020_Kundennummer0": docItem.SAP_x0020_Kundennummer,
              "OData__Comments":docItem.OData__Comments,
              "Customer0":docItem.Customer0,
              "ResponsibleKAMId":docItem.ResponsibleKAMId,
              "ResponsibleKAMStringId": docItem.ResponsibleKAMStringId,
              "Customer_x0020_Classification":docItem.Customer_x0020_Classification,
              "Project":docItem.Project,
              "Country":{
                "Label":docItem.Country.Label,
                "TermGuid":docItem.Country.TermGuid,
                "WssId":-1
             },

            }
          }
          else if(docItem.AMANN_x0020_Company!=null ){
             listItemPayload = {
              "ContentTypeId": contentTypeId,
              "Title": docItem.Title,
              "Test_x0020_Division":"Automotive",
              "SAP_x0020_Kundennummer0": docItem.SAP_x0020_Kundennummer,
              "OData__Comments":docItem.OData__Comments,
              "Customer0":docItem.Customer0,
              "ResponsibleKAMId":docItem.ResponsibleKAMId,
              "ResponsibleKAMStringId": docItem.ResponsibleKAMStringId,
              "Customer_x0020_Classification":docItem.Customer_x0020_Classification,
              "Project":docItem.Project,
              "AMANN_x0020_Company":{
                "Label":docItem.AMANN_x0020_Company.Label,
                "TermGuid":docItem.AMANN_x0020_Company.TermGuid,
                "WssId":-1
             }
            }
          }
          else{
            listItemPayload = {
              "ContentTypeId": contentTypeId,
              "Title": docItem.Title,
              "Test_x0020_Division":"Automotive",
              "SAP_x0020_Kundennummer0": docItem.SAP_x0020_Kundennummer,
              "OData__Comments":docItem.OData__Comments,
              "Customer0":docItem.Customer0,
              "ResponsibleKAMId":docItem.ResponsibleKAMId,
              "ResponsibleKAMStringId": docItem.ResponsibleKAMStringId,
              "Customer_x0020_Classification":docItem.Customer_x0020_Classification,
              "Project":docItem.Project
            }
          }
    //         if(docItem.Country!=null && docItem.AMANN_x0020_Company!=null){
    //             listItemPayload = {
    //           // "__metadata": { "type": "SP.Data.YourListNameListItem" },
    //           "ContentTypeId": contentTypeId,
    //           "Title": docItem.Title,
    //           "Test_x0020_Division":"Automotive",
    //      //     "SAP_x0020_Kundennummer": docItem.SAP_x0020_Kundennummer,
    //           "_Comments":docItem._Comments,
    //           "Customer":docItem.Customer0,
    //           "ResponsibleKAMId":docItem.ResponsibleKAMId,
    //           "ResponsibleKAMStringId": docItem.ResponsibleKAMStringId,
    //           "Customer_x0020_Classification":docItem.Customer_x0020_Classification,
    //           "Country":{
    //             "Label":docItem.Country.Label,
    //             "TermGuid":docItem.Country.TermGuid,
    //             "WssId":-1
    //             },
    //             "AMANN_x0020_Company":{
    //               "Label":docItem.AMANN_x0020_Company.Label,
    //               "TermGuid":docItem.AMANN_x0020_Company.TermGuid,
    //               "WssId":-1
    //           }

    //         }
              
    //         }else if(docItem.Country!=null ){
    //            listItemPayload = {
    //             // "__metadata": { "type": "SP.Data.YourListNameListItem" },
    //             "ContentTypeId": contentTypeId,
    //             "Title": docItem.Title,
    //             "Test_x0020_Division":"Automotive",
    //    //         "SAP_x0020_Kundennummer": docItem.SAP_x0020_Kundennummer,
    //             "_Comments":docItem._Comments,
    //             "Customer":docItem.Customer0,
    //             "ResponsibleKAMId":docItem.ResponsibleKAMId,
    //             "ResponsibleKAMStringId": docItem.ResponsibleKAMStringId,
    //             "Customer_x0020_Classification":docItem.Customer_x0020_Classification,
    //             "Country":{
    //               "Label":docItem.Country.Label,
    //               "TermGuid":docItem.Country.TermGuid,
    //               "WssId":-1
    //            },
  
    //           }
    //         }
    //         else if(docItem.AMANN_x0020_Company!=null ){
    //            listItemPayload = {
    //             "ContentTypeId": contentTypeId,
    //             "Title": docItem.Title,
    //             "Test_x0020_Division":"Automotive",
    //  //           "SAP_x0020_Kundennummer": docItem.SAP_x0020_Kundennummer,
    //             "_Comments":docItem._Comments,
    //             "Customer":docItem.Customer0,
    //             "ResponsibleKAMId":docItem.ResponsibleKAMId,
    //             "ResponsibleKAMStringId": docItem.ResponsibleKAMStringId,
    //             "Customer_x0020_Classification":docItem.Customer_x0020_Classification,
    //             "AMANN_x0020_Company":{
    //               "Label":docItem.AMANN_x0020_Company.Label,
    //               "TermGuid":docItem.AMANN_x0020_Company.TermGuid,
    //               "WssId":-1
    //            }
    //           }
    //         }
    //         else{
    //           listItemPayload = {
    //             "ContentTypeId": contentTypeId,
    //             "Title": docItem.Title,
    //             "Test_x0020_Division":"Automotive",
    //     //        "SAP_x0020_Kundennummer": docItem.SAP_x0020_Kundennummer,
    //             "_Comments":docItem._Comments,
    //             "Customer":docItem.Customer0,
    //             "ResponsibleKAMId":docItem.ResponsibleKAMId,
    //             "ResponsibleKAMStringId": docItem.ResponsibleKAMStringId,
    //             "Customer_x0020_Classification":docItem.Customer_x0020_Classification,
    //           }

    //         }


            console.log("listItemPayload:",JSON.stringify(listItemPayload));
            
            const listItemResponse = await this._webPartContext.spHttpClient.post(listItemUrl, SPHttpClient.configurations.v1, {
              headers: {
                "Accept": "application/json",
                "Content-Type": "application/json",
                "X-RequestDigest": this._webPartContext.pageContext.legacyPageContext.formDigestValue,
                "X-HTTP-Method": "MERGE",
                "If-Match": "*"
              },
              body: JSON.stringify(listItemPayload)
            });
      
            if (listItemResponse.ok) {
              return await listItemResponse.json().then((xtest=>{
                console.log("xtest",xtest);
              }))
            } else {
              console.log(`Error updating list item: ${listItemResponse.status}`);
              return [];
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
  public async updateUpload(itemId:string,docItem:ITodoItem, libraryName:string,documents: File[]):Promise<any> {
    try {
      let listUrl= '';
      if(itemId== null || itemId== '' || itemId== undefined){
        console.log(itemId, "item id is missing to update the file props")
        return null;
      }
      let urlreq =  `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${libraryName}')/items(${itemId})`;
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
      console.log("update dupload:",division, itemId);
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
      }

    }
    console.log("listItemPayload",listItemPayload);
     
      const listItemResponse = await this._webPartContext.spHttpClient.post(urlreq, SPHttpClient.configurations.v1, {
      headers: {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "X-RequestDigest": this._webPartContext.pageContext.legacyPageContext.formDigestValue,
          "X-HTTP-Method": "MERGE",
        "If-Match": "*"
      },
      body: JSON.stringify(listItemPayload)
      });
  
      if(listItemResponse.ok) {
        return await listItemResponse.json();
      }
      else {
        console.log(`Error updating list item: ${listItemResponse.status}`);
        return null;;
      }
      
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
  public async finalUploadV2(docItem:ITodoItem, libraryName:string,documents: File[],  contentTypeId:string):Promise<any> {
    console.log("new doc creating ct:",contentTypeId);
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


    var url = `${webUrl}/_api/web/lists/getByTitle('${libraryName}')/RootFolder/Files/Add(url='${fileName}', overwrite=true)`;
            // Construct the Endpoint  
      // var url = webUrl + "/_api/Web/GetFolderByServerRelativeUrl(@target)/Files/add(overwrite=true, url='" + fileName + "')?@target='" + targetUrl + "'&$expand=ListItemAllFields";  
    // Get the SharePoint list or library endpoint
    //const endpointUrl = `${webUrl}/_api/web/getfolderbyserverrelativeurl('${listUrl}')/files/add(overwrite=true,url='${fileName}')?`;

    // Convert the file content to an ArrayBuffer
    //const contentArrayBuffer = new TextEncoder().encode(fileContent);

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
        const filteredList =groups.filter(item =>  item.Title.toLowerCase().includes("member") ||   item.Title.toLowerCase().includes("owner") )[0];
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
  public  _getPermissions = async () => {
    try {
      const response = await fetch('/_api/web/effectiveBasePermissions', {
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
        },
      });

      if (response.ok) {
        const data = await response.json();
        //console.log(data);
        const permissions = data.d.EffectiveBasePermissions;
        let viewItems = permissions.High & 1 ? 'Yes' : 'No';
        let addtems = permissions.High & 2 ? 'Yes' : 'No';
        let editItems = permissions.High & 4 ? 'Yes' : 'No';
        let deleteItems = permissions.High & 8 ? 'Yes' : 'No';
        // console.log("viewItems",viewItems);
        // console.log("addtems",addtems);
        // console.log("editItems",editItems);
        // console.log("deleteItems",deleteItems);
        // setPermissions(permissions);
      } else {
        throw new Error('Failed to get user permissions');
      }
    } catch (error) {
      console.error(error);
    }
  };
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
    console.log("_getLists");
    const response: SPHttpClientResponse = await this._webPartContext.spHttpClient.get(endpointUrl, SPHttpClient.configurations.v1,options);
    return await response.json().then((json: { value: ITodoTaskList[] }) => {
      return  SharePointDataProvider._taskLists = json.value.map( (task: ITodoTaskList) => {
       let test:ITodoTaskList =task;
       console.log(test);
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
  private _getDocs(requester: SPHttpClient): Promise<ITodoItem[]> {
   
    var sortOrder = "desc";
    //const queryString: string =`?$select=Title,ID,Created,UniqueId,FileRef,FileDirRef,LinkFilename,LinkFilename2,ServerUrl,FileLeafRef,ContentTypeId,Author/Title,Editor/Title&$expand=Editor/Title&expand=Modified/ID,Modified/Title&$expand=Author/Title,File/LinkingUrl&$expand=File&$top=3000&$orderby=Created `+sortOrder;
   // const queryString: string =`?$select=Title,Responsible_x0020_CSC,ResponsibleKAM,Country,AMANN_x0020_Company,Priority,time_x0020_customer,Received,Incoterms_x0020__x0028_currently_x0029_,Terms_x0020_of_x0020_payment,Label0,Trunover_x0020__x0028_Prev_x002e_Year_x002d_YTD_x0029_,Comments,Customer_x0020_Classification,ID,Created,UniqueId,FileRef,FileDirRef,LinkFilename,LinkFilename2,ServerUrl,FileLeafRef,ContentTypeId,Author/Title,Editor/Title&$expand=Editor/Title&expand=Modified/ID,Modified/Title&$expand=Author/Title,File/LinkingUrl&$expand=File&$top=3000&$orderby=Created `+sortOrder;
    const queryString: string =`?$select=Title,Project,SAP_x0020_Kundennummer,OData__Comments,Responsible_x0020_CSC0Id,Responsible_x0020_CSC0StringId,Category_x0020_of_x0020_Document,Customer_x0020_Classification,Responsible_x0020_CSCId,Responsible_x0020_CSCStringId,ResponsibleKAMStringId,ResponsibleKAMId,Customer0,Turnover_x0020__x0028_Prev_x002e_Year_x0029_,Responible_x0020_Legal,Country,AMANN_x0020_Company,Priority,Deadline_x0020_Customer,Received0,Incoterms_x0020__x0028_current_x0029_,Label0,Trunover_x0020__x0028_Prev_x002e_Year_x002d_YTD_x0029_,Customer_x0020_Classification,ID,Created,UniqueId,FileRef,FileDirRef,LinkFilename,LinkFilename2,ServerUrl,FileLeafRef,ContentTypeId,Author/Title,Editor/Title&$expand=Editor/Title&expand=Modified/ID,Modified/Title&$expand=Author/Title,File/LinkingUrl&$expand=File&$top=3000&$orderby=Modified `+sortOrder;
    
    //const queryString: string =`?$select=*&$top=3000&$orderby=Created `+sortOrder;
    const queryUrl: string = this._listItemsUrl + queryString;
    // const requestOptions: ISPHttpClientOptions = {
    //   headers: {
    //     'Accept': 'application/octet-stream'
    //   }
    // };
    return requester.get(queryUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((json: { value: ITodoItem[] }) => {
       // console.log("GetDocs: ",json.value);
        return json.value.map((task: ITodoItem) => {
          task.DefaultEditUrl = `${this._webPartContext.pageContext.web.absoluteUrl}/_layouts/15/listform.aspx?PageType=6&ListId=${this.selectedList.Id}&ID=${task.Id}&RootFolder=*`;
         
          return task;
        });
      });
  }
  private _getDoc(itemId: string): Promise<ITodoItem> {
    console.log("itemId",itemId);
    const queryString: string =`?$select=Title,Project,SAP_x0020_Kundennummer,Category_x0020_of_x0020_Document,OData__Comments,Responsible_x0020_CSC0Id,Responsible_x0020_CSC0StringId,Customer_x0020_Classification,Responsible_x0020_CSCId,Responsible_x0020_CSCStringId,ResponsibleKAMStringId,ResponsibleKAMId,Customer0,Turnover_x0020__x0028_Prev_x002e_Year_x0029_,Responible_x0020_Legal,Country,AMANN_x0020_Company,Priority,Deadline_x0020_Customer,Received0,Incoterms_x0020__x0028_current_x0029_,Label0,Trunover_x0020__x0028_Prev_x002e_Year_x002d_YTD_x0029_,Customer_x0020_Classification,ID,Created,UniqueId,FileRef,FileDirRef,LinkFilename,LinkFilename2,ServerUrl,FileLeafRef,ContentTypeId,Author/Title,Editor/Title&$expand=Editor/Title&expand=Modified/ID,Modified/Title&$expand=Author/Title,File/LinkingUrl&$expand=File&$filter=Id eq '${itemId}'`;
    //const queryString: string =`?$select=*&$top=3000&$orderby=Created `+sortOrder;
    const queryUrl: string = this._listItemsUrl + queryString;
    return this._webPartContext.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json()
        .then((dataitm =>{
         // console.log("dataitm",dataitm);
          let itm:ITodoItem = dataitm.value[0];
         // console.log("itm,",itm);
          return itm;
        }))
      });
  }
  public createDoc(documents: File[]): Promise<ITodoItem[]> {
    return this
      ._createDoc(documents, this.webPartContext.spHttpClient)
      .then(_ => {
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
    //const body = new FormData();
    let url = this._docItemsUrl + `/RootFolder/Files/Add(url='${file.name}', overwrite=true)`;
        // body.append('@data.type', this._selectedList.EntityTypeName);
        // body.append('Title', file.name);
        // body.append('file', file);
        return client.post(
          url,
          SPHttpClient.configurations.v1,
          spOpts
        ); 
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
    console.log("convertedText",convertedText);
    
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
    // public async getCretedDocumentSetId(filteredList,contentTypeId,documentSetname,internalName){

  //   try {
  //     //documentSetname = documentSetname+"_docset";
  //     var encodedDocumentSetname = encodeURIComponent(documentSetname);
  //     let listUrl = `${this._webPartContext.pageContext.web.absoluteUrl}/${internalName}`;
  //     //const digestValue = await this.getRequestDigestValue();
  //     let slug =  `${listUrl}/${documentSetname}|${contentTypeId}` ;
  //   //   const folderData = {
  //   //     Title: documentSetname,
  //   //     Path: `${internalName}/${documentSetname}}`,
  //   //     ContentTypeId: contentTypeId,
  //   //     FileSystemObjectType: 1,
  //   //     ListItemEntityTypeFullName: filteredList.ListItemEntityTypeFullName
  //   // };

  //   const payload = {
  //     '__metadata': { 'type': 'SP.DocumentSet' },
  //     'ContentTypeId': contentTypeId,
  //     'Title':documentSetname,
  //     'Path': `${listUrl}`,
  //     'Name': documentSetname
  //   };

  //   // payload: JSON.stringify({
  //   //   '__metadata': { 'type': 'SP.DocumentSet' },
  //   //   'ContentTypeId': contentTypeId,
  //   //   'PropertyName1': 'New Value 1',
  //   //   'PropertyName2': 'New Value 2',
  //   //   // Add other properties to update
  //   // })

  //     let options: ISPHttpClientOptions =  { 
  //       // method: "POST", 
  //       // body: JSON.stringify(
  //       //   { 'Title' : documentSetname , 'Path' : listUrl }), 
  //       body: JSON.stringify(payload),


  //       headers: 
  //         { 
  //           "content-type": "application/json;odata=verbose", 
  //           "accept": "application/json;odata=verbose", 
  //           "slug":slug
  //         } 
  //       };
  //      let reqURL = `${this._webPartContext.pageContext.web.absoluteUrl}/_vti_bin/listdata.svc/${internalName}`;
  //    //let reqURL = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists('4010f6e6-3ee9-4740-a247-20afaa652f1d')/RootFolder/Files/add(url='test1234', overwrite=true)`;
  //   // let reqURL = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists('${filteredList.Id}')/items/add(url='test1234', overwrite=true)`;
     
     
  //    // let reqURL = `${this._webPartContext.pageContext.web.absoluteUrl}/_vti_bin/listdata.svc/CustomerDocumentsAutomotive`;
  //     const response = await this._webPartContext.spHttpClient.post(reqURL, SPHttpClient.configurations.v1,options)
  //     .then((response: SPHttpClientResponse) =>{
  //       if (response.ok){
  //         return response.json();
  //       }
  //       return response;
  //     }).catch((error: any) => {
  //       console.log(error);
  //     });
  //       // .then(  data =>{
  //       //   return data;
          

  //       // });

  //     return response;
  //   } catch (error) {
  //     console.log(error);
          
  //   }
  //   return null;
  // }

  //  public async moveFile(sourceFileUrl:string,destFileUrl:string ,contentTypeId:string):Promise<any> {
  //   try {
  //     // sourceFileUrl = "/sites/spclassicdev/DropLibrary/Banglalink UAT_ODA.docx";
  //     // destFileUrl = "/sites/spclassicdev/copied/Banglalink UAT_ODA/Banglalink UAT_ODA.docx";
  //    const webUrl = this._webPartContext.pageContext.web.absoluteUrl;
  //    const apiUrl = `${webUrl}/_api/web/getfilebyserverrelativeurl('${sourceFileUrl}')/moveto(newurl='${destFileUrl}',flags=1)`;
  //   // const apiUrl = `${webUrl}/_api/web/getfilebyserverrelativeurl('${fileUrl}')/copyto(strnewurl='${newFileUrl}',boverwrite=true)`;
  //    const headers = {
  //      "X-RequestDigest": this._webPartContext.pageContext.legacyPageContext.formDigestValue,
  //      "Accept": "application/json",
  //      "Content-Type": "application/json", // Use the MIME type of the file being uploaded
  //      "X-HTTP-Method": "POST",
  //      "X-Microsoft-HTTP-Method": "PUT",
  //      "If-Match": "*"
  //    };

  //   //  const body = JSON.stringify({
  //   //   "__metadata": {
  //   //     "type": "SP.ListItem"
  //   //   },
  //   //   "ContentTypeId": contentTypeId,
  //   // });
 
  //    const spHttpClientConfig = {
  //      headers: headers,
  //      body: ""
  //    };
 
  //    const spHttpClientOptions = {
  //      method: "POST",
  //      spHttpClientConfig: spHttpClientConfig
  //    };
 
  //    return await this.sendRequest(apiUrl,spHttpClientOptions)
  //      .then( async response => {
  //       console.log("moveFile",response)
  //        //console.log(response);

  //        console.log("File moved successfully");

  //        console.log("result:", x);
  //        const fileId =parseInt(documentSetId) +1;
  //        const listItemUrl = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${libraryName}')/items(${fileId})`;
  //        const listItemPayload = {
  //          // "__metadata": { "type": "SP.Data.YourListNameListItem" },
  //          "ContentTypeId": contentTypeId
  //        };
  //        const listItemResponse = await this._webPartContext.spHttpClient.post(listItemUrl, SPHttpClient.configurations.v1, {
  //          headers: {
  //            "Accept": "application/json",
  //            "Content-Type": "application/json",
  //            "X-RequestDigest": this._webPartContext.pageContext.legacyPageContext.formDigestValue,
  //            "X-HTTP-Method": "MERGE",
  //            "If-Match": "*"
  //          },
  //          body: JSON.stringify(listItemPayload)
  //        });
   
  //        if (listItemResponse.ok) {
  //          return await listItemResponse.json().then((xtest=>{
  //            console.log("xtest",xtest);
  //          }))
  //        } else {
  //          console.log(`Error updating list item: ${listItemResponse.status}`);
  //          return [];
  //        }
  //       //  console.log("result:", x);
  //       //     const fileId =parseInt(documentSetId) +1;

  //       // const itemUrl = `${webUrl}/_api/web/GetFileByServerRelativeUrl('${destFileUrl}')/ListItemAllFields`;
  //       // const properties = {
  //       //   'Title': 'test12344444',

  //       // };

  //       // const headers = {
  //       //   Accept: "application/json;odata=verbose",
  //       //   "Content-Type": "application/json;odata=verbose",
  //       //   "X-RequestDigest": this._webPartContext.pageContext.legacyPageContext.formDigestValue,
  //       //   "X-HTTP-Method": "MERGE",
  //       //   "IF-MATCH": "*"
  //       // };
      
  //       // const itemPayload = {
  //       //   // __metadata: {
  //       //   //   type: "SP.File"
  //       //   // }
  //       // };
      
  //       // for (const prop in properties) {
  //       //   itemPayload[prop] = properties[prop];
  //       // }
      
  //       // const spHttpClientOptions: ISPHttpClientOptions = {
  //       //   body: JSON.stringify(itemPayload),
  //       //   headers: headers
  //       // };
  //       // //GetFileByServerRelativeUrl('/sites/spclassicdev/Copied/Banglalink%20UAT_ODA/Banglalink%20UAT_ODA.docx')/ListItemAllFields

  //       //      const queryUrl = itemUrl;//`${this._webPartContext.pageContext.web.absoluteUrl}/_api/Web/GetFolderByServerRelativePath(decodedurl='/sites/spclassicdev/copied/Banglalink UAT_ODA')?$expand=Folders,Files,ListItemAllFields`;
  //       //      this._webPartContext.spHttpClient.post(queryUrl, SPHttpClient.configurations.v1,spHttpClientOptions)
  //       //      .then((response: SPHttpClientResponse) => {
  //       //       console.log("moveFile",response);
  //       //       response.json().then((x=>{
  //       //         console.log("movefile_v",x);
  //       //       }))
  //       //      })
             
  //       //     const listItemPayload = {
  //       //       // "__metadata": { "type": "SP.Data.YourListNameListItem" },
  //       //       "ContentTypeId": contentTypeId
  //       //     };
  //       //     const listItemResponse = await this._webPartContext.spHttpClient.post(listItemUrl, SPHttpClient.configurations.v1, {
  //       //       headers: {
  //       //         "Accept": "application/json",
  //       //         "Content-Type": "application/json",
  //       //         "X-RequestDigest": this._webPartContext.pageContext.legacyPageContext.formDigestValue,
  //       //         "X-HTTP-Method": "MERGE",
  //       //         "If-Match": "*"
  //       //       },
  //       //       body: JSON.stringify(listItemPayload)
  //       //     });
      
  //       //     if (listItemResponse.ok) {
  //       //       return await listItemResponse.json().then((xtest=>{
  //       //         console.log("xtest",xtest);
  //       //       }))
  //       //     } else {
  //       //       console.log(`Error updating list item: ${listItemResponse.status}`);
  //       //       return [];
  //       //     }





  //        return response;
  //      })
    
  //      .catch(error => {
  //        console.log(error);
  //        console.log("error moving file");
  //        return error;
  //      });
       
  //   }
  //   catch (error) {
  //     console.log(error);
  //     console.log("error in catch moving file");
  //   }
  // }
  
  // public async sendRequest(apiUrl,options){
  //     //const spHttpClient = this._webPartContext.spHttpClient;
  //     return await this._webPartContext.spHttpClient.fetch(apiUrl,SPHttpClient.configurations.v1,options)
  //       .then(response => {
  //         console.log("response.ok",response.ok);
  //         if (!response.ok) {
  //            return  response;
  //         }else{
  //           console.log("saved");
  //           return response;
  //         }
  //         //resolve(response.json());
  //       })
  //       .catch(error => {
  //         console.log("log catch error");
  //         //reject(error);
  //       });
  //   };


    // public async getRequiredFields(): Promise<any>{
  //   try {
  //     // const endpoint = `${this._docItemsUrl}/fields?$filter=Required%20eq%20true`;
  //     // const endpoint = `${this._docItemsUrl}/fields?$filter=Hidden eq 'false'`;

  //     const endpoint = `${this._docItemsUrl}/fields`;
  //     const response = await this._webPartContext.spHttpClient.get(endpoint, SPHttpClient.configurations.v1)
  //     .then((response: SPHttpClientResponse) =>{
  //       if (response.ok){
  //         return response.json();
  //       }});
  //     return response;
  //   } catch (error) {
  //     console.log(error);
  //     return [];
  //   }
  // }
  // public async moveFileByPath(srcPath: string, destPath: string, name: string, shouldOverWrite: boolean, keepBoth: boolean): Promise<void> {
  //    srcPath = `/sites/spclassicdev/DropLibrary/${name}`;
  //    destPath =`/sites/spclassicdev/copied/data`;
    
  //   //const endpoint = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/getFileByServerRelativePath('${srcPath}')/MoveTo(newPath='${destPath}/${name}',flags=${shouldOverWrite ? '1' : '0'},keepBoth=${keepBoth})`;
  //   const endpoint = `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/getFileByServerRelativePath('${encodeURIComponent(srcPath)}')/MoveTo(newPath='${encodeURIComponent(destPath)}/${encodeURIComponent(name)}',flags=${shouldOverWrite ? '1' : '0'},keepBoth=${keepBoth})`;
  //   try {
  //     const response: SPHttpClientResponse = await this._webPartContext.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {});
  //     if (response.ok) {
  //       console.log('File moved successfully');
  //     } else {
  //       console.log('File move failed');
  //     }
  //   } catch (error) {
  //     console.error('Error moving file:', error);
  //   }
  // }
}
