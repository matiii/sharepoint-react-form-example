import ListItemService from './ListItemService';
import PeopleService from './PeopleService';

import ListItemServiceMock from './mock/ListItemService';
import PeopleServiceMock from './mock/PeopleService';

let environment = 'dev';

function isProductionEnvironment(): boolean {
    return environment !== 'dev';
}

export function getListItemService(): typeof ListItemService {
    return isProductionEnvironment() ? ListItemService : ListItemServiceMock;
}

export function getPeopleService(): typeof PeopleService {
    return isProductionEnvironment() ? PeopleService : PeopleServiceMock;
}