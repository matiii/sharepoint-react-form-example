import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';

export interface IListItemService {
    setItem(listName: string, data: { property: string, value: any }[], success: (listItemId: number) => void, error: (msg: string) => void);

    addAttachment(listName: string, listItemId: number, fileName: string, base64: string, success: Function, error: (reason: any) => void);
}

export interface IPeopleService {
    getPeoplesByString(text: string, maxResults: number): Promise<(IPersonaProps & { key: string | number })[]>;

}