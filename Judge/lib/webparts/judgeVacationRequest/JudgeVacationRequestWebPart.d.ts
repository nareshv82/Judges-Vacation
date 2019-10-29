import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration } from '@microsoft/sp-webpart-base';
export interface IJudgeVacationRequestWebPartProps {
    ListName: string;
}
export default class JudgeVacationRequestWebPart extends BaseClientSideWebPart<IJudgeVacationRequestWebPartProps> {
    private listItemEntityTypeName;
    private Department;
    private Manager;
    protected onInit(): Promise<void>;
    constructor();
    render(): void;
    private setButtonsState;
    private getSPData;
    private renderData;
    private formatDate;
    private setButtonsEventHandlers;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
    private listNotConfigured;
    private createItem;
    private readItem;
    private getLatestItemId;
    private readItems;
    private updateItem;
    private deleteItem;
    private updateStatus;
    private updateItemsHtml;
}
//# sourceMappingURL=JudgeVacationRequestWebPart.d.ts.map