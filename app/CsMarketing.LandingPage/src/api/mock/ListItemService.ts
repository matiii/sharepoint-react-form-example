export default class ListItemService {

    setItem(listName: string, data: { property: string, value: any }[], success: (listItemId: number) => void, error: (msg: string) => void) {

        success(1);

    }

    addAttachment(listName: string,
        listItemId: number,
        fileName: string,
        base64: string,
        success: Function,
        error: (reason: any) => void) {

        success();

    }

}