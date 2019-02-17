import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GetListInfoWebPart.module.scss';
import * as strings from 'GetListInfoWebPartStrings';

//ヘルパークラス
import{SPHttpClient,SPHttpClientResponse} from '@microsoft/sp-http';

//ISPList
import {ISPLists, ISPList} from './ISPList';

//環境の切り替え
import {Environment, EnvironmentType} from '@microsoft/sp-core-library';

export interface IGetListInfoWebPartProps {
  description: string;
}

export default class GetListInfoWebPart extends BaseClientSideWebPart<IGetListInfoWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.getListInfo }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">
              ${escape(this.context.pageContext.web.title)}サイトのリスト一覧
              </p>
              </a>
            </div>
          </div>
        </div>
        <div id="spListContainer"/>
      </div>`;
      this._renderListAsync();
  }

  //SharePoint REST APIの呼び出し
  private _getListData():Promise<ISPLists>{
    return this.context.spHttpClient.get(
      this.context.pageContext.web.absoluteUrl+
      `/_api/web/lists?$select=Title,Id,Description,Created,ImagePath/DecodedUrl,RootFolder/ServerRelativeUrl
      &$filter=Hidden eq false and BaseType eq 1
      &$expand=RootFolder`,
      SPHttpClient.configurations.v1)
      .then((Response:SPHttpClientResponse)=>{
        return Response.json();
      });
  }

  //環境を切り替えるためのメソッド
  private _renderListAsync():void{
    if(Environment.type==EnvironmentType.SharePoint ||
     Environment.type==EnvironmentType.ClassicSharePoint){
       this._getListData()
        .then((response)=>{
          this._renderList(response.value);
        });
     }
  }

  //REST API呼び出し結果をレンダリングするメソッド
  private _renderList(items:ISPList[]):void{
    let html:string='';
    items.forEach((item:ISPList)=>{
      let createdDate:Date=new Date(item.Created);
      html+=`
        <ul class="${styles.list}">
          <li class="${styles.listItem}">
          <a href="https://${window.location.hostname}${item.RootFolder.ServerRelativeUrl}"
          target="_blank" rel="noopener noreferer">
          <h1>
          <img src="${item.ImagePath.DecodedUrl}">
          ${item.Title}
          </h1></a>
          <p>${item.Description}</p>
          <hr>
          <p class="created">作成日:${createdDate.toLocaleDateString("ja-JP")}
          | GUID : ${item.Id}</p>
          </li>
        </ul>`;
    });

    const listContainer : Element = this.domElement.querySelector("#spListContainer");
    listContainer.innerHTML=html;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }


}
