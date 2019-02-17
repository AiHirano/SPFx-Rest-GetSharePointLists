export interface ISPLists{
    value:ISPList[];
}

export interface ISPList{
    Title:string;
    Id:string;
    Description:string;
    ImagePath:{DecodedUrl:string;};
    RootFolder:{ServerRelativeUrl:string;};
    Created:string;
}