interface ITodoItem {
  Id: number;
  Title: string;
  Url: string;
  LinkFilename:string;
  Author:string;
  FileLeafRef:string;
  Modified:string;
  Editor:string;
  Created:string;
  ContentTypeId:string;
  Division:string;
  ServerUrl:string;
  File:File;
  DefaultEditUrl:string;
  DefaultDisplayUrl:string;
  SAP_x0020_Kundennummer: string;
  _Comments:string;
  Customer_x0020_Classification:string;
  Label0: string;
  Customer0:string;
  Turnover_x0020__x0028_Prev_x002e_Year_x0029_:string;
  StartDate:string;
  Responible_x0020_Legal:string;
  Country:{
    Label:string;
    TermGuid:string;
    WssId:string;
  };
  AMANN_x0020_Company:{
    Label:string;
    TermGuid:string;
    WssId:string;
  };
  OData__Comments:string;
  Responsible_x0020_CSCStringId:string;
  Responsible_x0020_CSCId:string;

  Responsible_x0020_CSC0Id:string;
  Responsible_x0020_CSC0StringId:string;
  ResponsibleKAMStringId:string;
  ResponsibleKAMId :string;
  Category_x0020_of_x0020_Document:string;
  Project:string;
  Terms_x0020_of_x0020_Payment:string;
  

}

export default ITodoItem;
