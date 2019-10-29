import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import {SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions}  from '@microsoft/sp-http';
import {IListItem} from './loc/IListItem';
import { sp, Item, ItemAddResult, ItemUpdateResult, Site } from '@pnp/sp';
import {SPComponentLoader} from "@microsoft/sp-loader";
import { CurrentUser } from '@pnp/sp/src/siteusers';

import styles from './JudgeVacationRequestWebPart.module.scss';
import * as strings from 'JudgeVacationRequestWebPartStrings';

export interface IJudgeVacationRequestWebPartProps {
  ListName: string;
}

export default class JudgeVacationRequestWebPart extends BaseClientSideWebPart<IJudgeVacationRequestWebPartProps> {
  private listItemEntityTypeName: string = undefined;
  private Department: string;
  private Manager:Int32Array;
  protected onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error?: any) => void): void => {
      sp.setup({
        sp: {
          headers: {
            "Accept": "application/json; odata=nometadata"
          }
        }
      });
      resolve();
    });
  }
  public constructor(){
    super();
      SPComponentLoader.loadScript('https://code.jquery.com/jquery-1.12.4.min.js',{}).then(()=>{
      SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css');
      SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css');
      SPComponentLoader.loadScript('https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js');
      SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.6.4/css/bootstrap-datepicker.min.css'); 
      SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.6.4/js/bootstrap-datepicker.min.js');
    });
  }
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.judgeVacationRequest }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.ListName)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
          <div class="row">
            <div class="col-md-12">
                <h1>Judges timekeeping solution</h1>
                <ul>
                    <li>
                        <a href="#" data-toggle="modal" data-target="#JudgeVacationModal" onclick="return false">Judges Vacation Request Form</a>
                    </li>                  
                </ul>
            </div>
          </div>
          <div class="modal fade in" id="JudgeVacationModal" role="dialog">
          <div class="modal-dialog modal-xl" role="document">
              <div class="modal-content">
                  <div class="form-wrap">
                      <div class="modal-header">
                          <h4 class="modal-title thin">Judge Vacation Request</h4>
                          <button type="button" class="close" data-dismiss="modal" aria-label="Close" style="margin-top:-25px;margin-right:-5px;">
                              <span aria-hidden="true">×</span>
                          </button>
                      </div>
                      <div class="modal-body">
                          <div class="JudgeVacationForm">
                              <div class="row">
                                  <div class="col-md-12">
                                      <div class="row">
                                          <div class="col-md-12">
                                              <div class="form-group">
                                                  <label class="control-label">Requester:</label>
                                                  <div class="input-group">
                                                      <span class="input-group-addon">
                                                          <i class="fa fa-user" aria-hidden="true"></i>
                                                      </span>
                                                      <input type="text" class="form-control" placeholder="Enter Requester Name" data-info="RequesterName" data-error="Enter a Requester Name!" required="" id="requesterName">
                                                  </div>
                                                  <span class="glyphicon form-control-feedback" aria-hidden="true"></span>
                                                  <div class="help-block with-errors"></div>
                                              </div>
                                          </div>                                    
                                      </div>

                                      <div class="row">
                                      <div class="col-sm-6">
                                      <div class="input-group date" data-provide="datepicker">
                                          <input type="text" class="form-control"  placeholder="Enter a Start Date"  id="startDate">
                                          <div class="input-group-addon">
                                              <span class="glyphicon glyphicon-calendar"></span>
                                          </div>
                                      </div>
                                  </div>
                                          <div class="col-md-12">
                                              <div class="form-group">
                                                  <label class="control-label">Request Start Date:</label>
                                                  <div class="input-group date" data-provide="datepicker">
                                                      <input type="text" class="form-control"   id="startDate">
                                                      <div class="input-group-addon">
                                                        <span class="glyphicon glyphicon-calendar"></span>
                                                      </div>
                                                  </div>
                                                  <span class="glyphicon form-control-feedback" style="margin-right:40px;" aria-hidden="true"></span>
                                                  <div class="help-block with-errors"></div>
                                              </div>
                                          </div>
                                      </div>
                                        <div class="row">
                                          <div class="col-md-12">
                                              <div class="form-group">
                                                  <label class="control-label">Request End Date:</label>
                                                  <div class="input-group date" data-provide="datepicker">
                                                      <input type="text" class="form-control" placeholder="Enter a End Date" id="endDate">
                                                      <span class="input-group-addon"><span class="glyphicon glyphicon-calendar"></span></span>
                                                  </div>
                                                  <span class="glyphicon form-control-feedback" style="margin-right:40px;" aria-hidden="true"></span>
                                                  <div class="help-block with-errors"></div>
                                              </div>
                                          </div>
                                      </div>
                                      <div class="row">
                                          <div class="col-md-12">                                        
                                              <div class="form-group has-feedback">
                                                  <label class="control-label">Amount of Time:</label>
                                                  <select class="form-control select2-hidden-accessible" data-info="AmountofTime" required="" tabindex="-1" aria-hidden="true" id="dayOff">
                                                      <option value="Full day">Full day</option>
                                                      <option value="Half day am">Half day am</option>
                                                      <option value="Half day pm">Half day pm</option>
                                                  </select>											
                                                  <span class="glyphicon form-control-feedback shiftLeft" aria-hidden="true"></span>
                                                  <div class="help-block with-errors"></div>
                                              </div>
                                          </div>                                    
                                      </div>    
                                      <div class="row">
                                          <div class="col-md-12">                                        
                                              <div class="form-group has-feedback">
                                                  <label class="control-label">Request Type:</label>
                                                  <select class="form-control select2-hidden-accessible" data-info="RequestType" required="" tabindex="-1" aria-hidden="true" id="requestType">
                                                      <option value="Vacation">Vacation</option>
                                                      <option value="Sick">Sick</option>
                                                  </select>											
                                                  <span class="glyphicon form-control-feedback shiftLeft" aria-hidden="true"></span>
                                                  <div class="help-block with-errors"></div>
                                              </div>
                                          </div>                                    
                                      </div>								
                                      <div class="row">
                                          <div class="col-md-12">
                                              <div class="form-group has-feedback">
                                                  <label class="control-label">Comments:</label>
                                                  <div class="input-group">
                                                      <span class="input-group-addon">
                                                          <i class="fa fa-pencil"></i>
                                                      </span>
                                                      <textarea class="form-control" rows="3" placeholder="Enter Comments..." data-info="Comments" id="requesterComments"></textarea>
                                                  </div>
                                                  <span class="glyphicon form-control-feedback" aria-hidden="true"></span>
                                                  <div class="help-block with-errors"></div>
                                              </div>
                                          </div>
                                      </div>

                                  </div>
                              </div>
                          </div>
                      </div>
                      <div class="modal-footer">
                          <div class="form-group">
                              <button type="submit" class="btn btn-primary create-Button"> <i class="fa fa-check" aria-hidden="true"></i>
                                  Submit</button>
                              <button class="btn btn-default" data-dismiss="modal"> <i class="fa fa-times" aria-hidden="true"></i> Cancel</button>
                          </div>
                          <div class="status"></div>
                          <div class="items"></div>
                      </div>
                      
                  </div>
              </div>
          </div>
      </div>
      <script type="text/javascritp">
      $(document).ready(function(){
        $('#startDate').datepicker();
        $('#endDate').datepicker();
      });
      </script>
        </div>
      </div>`;
      this.updateStatus(this.listNotConfigured() ? 'Please configure list in Web Part properties' : 'Ready');
      this.setButtonsState();
      this.setButtonsEventHandlers();
      this.getSPData();
  }
  private setButtonsState(): void {
    const buttons: NodeListOf<Element> = this.domElement.querySelectorAll(`button.${styles.button}`);
    const listNotConfigured: boolean = this.listNotConfigured();

    for (let i: number = 0; i < buttons.length; i++) {
      const button: Element = buttons.item(i);
      if (listNotConfigured) {
        button.setAttribute('disabled', 'disabled');
      }
      else {
        button.removeAttribute('disabled');
      }
    }
  }
  private getSPData(): void {    
    sp.profiles.myProperties.get().then((r: CurrentUser) => {
      console.log(r);
      let dept = r["UserProfileProperties"][11]["Value"];
      let mgr = r["UserProfileProperties"][14]["Value"];
      let mgrID;
      const payload: string = JSON.stringify({
        'logonName': mgr //this.context.pageContext.user.loginName // i:0#.f|membership|firstname.lastname@contoso.onmicrosoft.com      
      });      
      var postData: ISPHttpClientOptions = {
        body: payload
      };
      var endPoint = `${this.context.pageContext.site.absoluteUrl}/_api/web/ensureuser`;
      this.context.spHttpClient.post(endPoint,
        SPHttpClient.configurations.v1,
        postData)
        .then((response: SPHttpClientResponse) => {
          response.json().then((resposneJSON:any)=>{
            console.log("manager");
            console.log(resposneJSON);
            this.Manager=resposneJSON.Id;
          });
          
          //return response.json();
      });
      this.renderData(r['DisplayName'],dept,mgrID);
    });
  }
   
  private renderData(strResponse: string,strDept: string,strMgr: Int32Array): void {
    document.getElementById("requesterName")["value"] = strResponse;
    document.getElementById("startDate")["value"] = this.formatDate(new Date().toString());
    document.getElementById("endDate")["value"] = this.formatDate(new Date().toString());
    this.Department=strDept;
    //this.Manager=strMgr;
  }
  
  private formatDate = (date: string) => {
    return new Intl.DateTimeFormat('en-US', { 
        year: 'numeric',
        month: 'numeric',
        day: 'numeric' })
      .format(new Date(date));
}
  private setButtonsEventHandlers(): void {
    const webPart: JudgeVacationRequestWebPart = this;
    this.domElement.querySelector('button.create-Button').addEventListener('click', () => { webPart.createItem(); });
    //this.domElement.querySelector('button.read-Button').addEventListener('click', () => { webPart.readItem(); });
    //this.domElement.querySelector('button.readall-Button').addEventListener('click', () => { webPart.readItems(); });
    //this.domElement.querySelector('button.update-Button').addEventListener('click', () => { webPart.updateItem(); });
    //this.domElement.querySelector('button.delete-Button').addEventListener('click', () => { webPart.deleteItem(); });
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
                  label: strings.ListNameFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private listNotConfigured(): boolean {
    return this.properties.ListName === undefined ||
      this.properties.ListName === null ||
      this.properties.ListName.length === 0;
  }
  private createItem(): void {
    this.updateStatus('Creating item...');
    sp.web.currentUser.get().then(res=>{
      console.log(res.Id);
      sp.web.lists.getByTitle("Judges Vacation Calendar").items.add({
      
        'Title': "Out of Office - "+document.getElementById("requesterName")["value"],
        'RequesterId':res.Id,
        'StartDate': new Date(document.getElementById("startDate")["value"]),
        'EndDate': new Date(document.getElementById("endDate")["value"]+" 11:45:00 PM"),
        //'HalfDay': document.getElementById("dayOff")["value"],
        //'Vacation_x0020_Type':document.getElementById("requestType")["value"],
        //'Description':document.getElementById("requesterComments"),
        'Department':this.Department,
        'ManagerId':this.Manager
      }).then((result: ItemAddResult): void => {
        const item: IListItem = result.data as IListItem;
        this.updateStatus(`Item '${item.Title}' (ID: ${item.Id}) successfully created`);
      }, (error: any): void => {
        this.updateStatus('Error while creating the item: ' + error);
      });
    });
    
  }

  private readItem(): void {
    this.updateStatus('Loading latest items...');
    this.getLatestItemId()
      .then((itemId: number): Promise<IListItem> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        this.updateStatus(`Loading information about item ID: ${itemId}...`);
        return sp.web.lists.getByTitle(this.properties.ListName)
          .items.getById(itemId).select('Title', 'Id').get();
      })
      .then((item: IListItem): void => {
        this.updateStatus(`Item ID: ${item.Id}, Title: ${item.Title}`);
      }, (error: any): void => {
        this.updateStatus('Loading latest item failed with error: ' + error);
      });
  }

  private getLatestItemId(): Promise<number> {
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
      sp.web.lists.getByTitle(this.properties.ListName)
        .items.orderBy('Id', false).top(1).select('Id').get()
        .then((items: { Id: number }[]): void => {
          if (items.length === 0) {
            resolve(-1);
          }
          else {
            resolve(items[0].Id);
          }
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  private readItems(): void {
    this.updateStatus('Loading all items...');
    sp.web.lists.getByTitle(this.properties.ListName)
      .items.select('Title', 'Id').get()
      .then((items: IListItem[]): void => {
        this.updateStatus(`Successfully loaded ${items.length} items`, items);
      }, (error: any): void => {
        this.updateStatus('Loading all items failed with error: ' + error);
      });
  }

  private updateItem(): void {
    this.updateStatus('Loading latest items...');
    let latestItemId: number = undefined;
    let etag: string = undefined;

    this.getLatestItemId()
      .then((itemId: number): Promise<Item> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        latestItemId = itemId;
        this.updateStatus(`Loading information about item ID: ${itemId}...`);
        return sp.web.lists.getByTitle(this.properties.ListName)
          .items.getById(itemId).get(undefined, {
            headers: {
              'Accept': 'application/json;odata=minimalmetadata'
            }
          });
      })
      .then((item: Item): Promise<IListItem> => {
        etag = item["odata.etag"];
        return Promise.resolve((item as any) as IListItem);
      })
      .then((item: IListItem): Promise<ItemUpdateResult> => {
        return sp.web.lists.getByTitle(this.properties.ListName)
          .items.getById(item.Id).update({
            'Title': `Item ${new Date()}`
          }, etag);
      })
      .then((result: ItemUpdateResult): void => {
        this.updateStatus(`Item with ID: ${latestItemId} successfully updated`);
      }, (error: any): void => {
        this.updateStatus('Loading latest item failed with error: ' + error);
      });
  }

  private deleteItem(): void {
    if (!window.confirm('Are you sure you want to delete the latest item?')) {
      return;
    }

    this.updateStatus('Loading latest items...');
    let latestItemId: number = undefined;
    let etag: string = undefined;
    this.getLatestItemId()
      .then((itemId: number): Promise<Item> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        latestItemId = itemId;
        this.updateStatus(`Loading information about item ID: ${latestItemId}...`);
        return sp.web.lists.getByTitle(this.properties.ListName)
          .items.getById(latestItemId).select('Id').get(undefined, {
            headers: {
              'Accept': 'application/json;odata=minimalmetadata'
            }
          });
      })
      .then((item: Item): Promise<IListItem> => {
        etag = item["odata.etag"];
        return Promise.resolve((item as any) as IListItem);
      })
      .then((item: IListItem): Promise<void> => {
        this.updateStatus(`Deleting item with ID: ${latestItemId}...`);
        return sp.web.lists.getByTitle(this.properties.ListName)
          .items.getById(item.Id).delete(etag);
      })
      .then((): void => {
        this.updateStatus(`Item with ID: ${latestItemId} successfully deleted`);
      }, (error: any): void => {
        this.updateStatus(`Error deleting item: ${error}`);
      });
  }

  private updateStatus(status: string, items: IListItem[] = []): void {
    this.domElement.querySelector('.status').innerHTML = status;
    this.updateItemsHtml(items);
  }

  private updateItemsHtml(items: IListItem[]): void {
    this.domElement.querySelector('.items').innerHTML = items.map(item => `<li>${item.Title} (${item.Id})</li>`).join("");
  }
}
