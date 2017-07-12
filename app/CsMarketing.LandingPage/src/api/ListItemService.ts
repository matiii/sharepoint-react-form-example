/// <reference path="../../node_modules/@types/sharepoint/index.d.ts" />

import axios from 'axios';
import * as Endpoints from './Endpoints';

const addAttachmentEnvelope = (listName: string, listItemId: number, fileName: string, base64: string) => {
    return (
        `<?xml version="1.0" encoding="utf-8"?>
         <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
         <soap:Body>
         <AddAttachment xmlns="http://schemas.microsoft.com/sharepoint/soap/">
          <listName>${listName}</listName>
          <listItemID>${listItemId}</listItemID>
          <fileName>${fileName}</fileName>
          <attachment>${base64}</attachment>
        </AddAttachment>
        </soap:Body>
        </soap:Envelope>`
    );
}

export default class ListItemService {

    setItem(listName: string, data: { property: string, value: any }[], success: (listItemId: number) => void, error: (msg: string) => void) {
        var clientContext = SP.ClientContext.get_current();
        var oWebsite = clientContext.get_web();
        var oList = oWebsite.get_lists().getByTitle(listName);

        var itemCreateInfo = new SP.ListItemCreationInformation();
        var oListItem = oList.addItem(itemCreateInfo);

        for (let d of data) {
            oListItem.set_item(d.property, d.value);
        }

        oListItem.update();

        clientContext.load(oListItem);
        clientContext.executeQueryAsync(
            successHandler,
            errorHandler
        );

        function successHandler() {
            success(oListItem.get_id());
        }

        function errorHandler(sender: any, args: SP.ClientRequestFailedEventArgs) {
            error(args.get_message());
        }
    }

    addAttachment(listName: string, listItemId: number, fileName: string, base64: string, success: Function, error: (reason: any) => void) {

        axios.post(Endpoints.ListsEndpoint,
            addAttachmentEnvelope(listName, listItemId, fileName, base64),
            {
                headers: { 'Content-Type': 'application/soap+xml; charset=utf-8' }
            }).then(() => {
                success();
            }).catch(error);

    }

}