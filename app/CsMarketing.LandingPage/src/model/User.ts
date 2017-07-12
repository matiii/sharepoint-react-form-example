import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';

export interface User {
    userId?: number;
    loginName?: string;
    title?: string;
    email?: string;
}

export interface PersonaUser extends User, IPersonaProps {
    key: number | string;
}