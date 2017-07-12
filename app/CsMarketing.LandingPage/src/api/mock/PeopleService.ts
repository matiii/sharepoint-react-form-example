import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { Promise } from 'es6-promise';
import { User, PersonaUser } from '../../model/User';

export default class PeopleService {
    private readonly people: (IPersonaProps & { key: string | number } & { email: string})[] = [
        {
            key: 0,
            imageUrl: '',
            imageInitials: 'PV',
            primaryText: 'Annie Lindqvist',
            secondaryText: 'Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'annie.lindqvist@intive.com'
        },
        {
            key: 1,
            imageUrl: '',
            imageInitials: 'AR',
            primaryText: 'Aaron Reid',
            secondaryText: 'Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'aaron.reid@intive.com'
            
        },
        {
            key: 2,
            imageUrl: '',
            imageInitials: 'AL',
            primaryText: 'Alex Lundberg',
            secondaryText: 'Software Developer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'alex.lundberg@intive.com'
        },
        {
            key: 3,
            imageUrl: '',
            imageInitials: 'RK',
            primaryText: 'Roko Kolar',
            secondaryText: 'Financial Analyst',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'roko.kolar@intive.com'
        },
        {
            key: 4,
            imageUrl: '',
            imageInitials: 'CB',
            primaryText: 'Christian Bergqvist',
            secondaryText: 'Sr. Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'christian.bergqvist@intive.com'
            
        },
        {
            key: 5,
            imageUrl: '',
            imageInitials: 'VL',
            primaryText: 'Valentina Lovric',
            secondaryText: 'Design Developer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'valentinal.lovric@intive.com'
        },
        {
            key: 6,
            imageUrl: '',
            imageInitials: 'MS',
            primaryText: 'Maor Sharett',
            secondaryText: 'UX Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'maor.sharett@intive.com'
            
        },
        {
            key: 7,
            imageUrl: '',
            imageInitials: 'PV',
            primaryText: 'Anny Lindqvist',
            secondaryText: 'Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'anny.lindqvist@intive.com'
        },
        {
            key: 8,
            imageUrl: '',
            imageInitials: 'AR',
            primaryText: 'Aron Reid',
            secondaryText: 'Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'aron.reid@intive.com'
        },
        {
            key: 9,
            imageUrl: '',
            imageInitials: 'AL',
            primaryText: 'Alix Lundberg',
            secondaryText: 'Software Developer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'alix.lundberg@intive.com'
        },
        {
            key: 10,
            imageUrl: '',
            imageInitials: 'RK',
            primaryText: 'Roko Kular',
            secondaryText: 'Financial Analyst',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'roko.kular@intive.com'
        },
        {
            key: 11,
            imageUrl: '',
            imageInitials: 'CB',
            primaryText: 'Christian Bergqvest',
            secondaryText: 'Sr. Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'christian.bergqvest@intive.com'
        },
        {
            key: 12,
            imageUrl: '',
            imageInitials: 'VL',
            primaryText: 'Valintina Lovric',
            secondaryText: 'Design Developer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'valintina.lovric@intive.com'
        },
        {
            key: 13,
            imageUrl: '',
            imageInitials: 'MS',
            primaryText: 'Maor Sharet',
            secondaryText: 'UX Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'maor.sharet@intive.com'
        },
        {
            key: 14,
            imageUrl: '',
            imageInitials: 'VL',
            primaryText: 'Anny Lindqvest',
            secondaryText: 'SDE',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'annie.lindqvist@intive.com'
            
        },
        {
            key: 15,
            imageUrl: '',
            imageInitials: 'MS',
            primaryText: 'Alix Lunberg',
            secondaryText: 'SE',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'alix.lunberg@intive.com'
        },
        {
            key: 16,
            imageUrl: '',
            imageInitials: 'VL',
            primaryText: 'Annie Lindqvest',
            secondaryText: 'SDET',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'annie.lindqvist@intive.com'
            
        },
        {
            key: 17,
            imageUrl: '',
            imageInitials: 'MS',
            primaryText: 'Alixander Lundberg',
            secondaryText: 'Senior Manager of SDET',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'alixander.lundberg@intive.com'
        },
        {
            key: 18,
            imageUrl: '',
            imageInitials: 'VL',
            primaryText: 'Anny Lundqvist',
            secondaryText: 'Junior Manager of Software',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'annie.lindqvist@intive.com'
        },
        {
            key: 19,
            imageUrl: '',
            imageInitials: 'MS',
            primaryText: 'Maor Shorett',
            secondaryText: 'UX Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'maor.shorett@intive.com'
        },
        {
            key: 20,
            imageUrl: '',
            imageInitials: 'VL',
            primaryText: 'Valentina Lovrics',
            secondaryText: 'Design Developer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'valentina.lovrics@intive.com'
        },
        {
            key: 21,
            imageUrl: '',
            imageInitials: 'MS',
            primaryText: 'Maor Sharet',
            secondaryText: 'UX Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'maor.sharet@intive.com'
        },
        {
            key: 22,
            imageUrl: '',
            imageInitials: 'VL',
            primaryText: 'Valentina Lovrecs',
            secondaryText: 'Design Developer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'valentina.lovrecs@intive.com'
        },
        {
            key: 23,
            imageUrl: '',
            imageInitials: 'MS',
            primaryText: 'Maor Sharitt',
            secondaryText: 'UX Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'maor.sharitt@intive.com'
        },
        {
            key: 24,
            imageUrl: '',
            imageInitials: 'MS',
            primaryText: 'Maor Shariett',
            secondaryText: 'Design Developer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 3:00pm',
            email: 'maor.sharitt@intive.com'
            
        },
        {
            key: 25,
            imageUrl: '',
            imageInitials: 'AL',
            primaryText: 'Alix Lundburg',
            secondaryText: 'UX Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 3:00pm',
            email: 'alix.lundburg@intive.com'
        },
        {
            key: 26,
            imageUrl: '',
            imageInitials: 'VL',
            primaryText: 'Valantena Lovric',
            secondaryText: 'UX Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'valantena.lovric@intive.com'
        },
        {
            key: 27,
            imageUrl: '',
            imageInitials: 'VL',
            primaryText: 'Velatine Lourvric',
            secondaryText: 'UX Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'valantena.lovric@intive.com'
        },
        {
            key: 28,
            imageUrl: '',
            imageInitials: 'VL',
            primaryText: 'Valentyna Lovrique',
            secondaryText: 'UX Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'valantena.lovric@intive.com'
        },
        {
            key: 29,
            imageUrl: '',
            imageInitials: 'AL',
            primaryText: 'Annie Lindquest',
            secondaryText: 'UX Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'annie.lindqvist@intive.com'
        },
        {
            key: 30,
            imageUrl: '',
            imageInitials: 'AL',
            primaryText: 'Anne Lindquist',
            secondaryText: 'UX Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'annie.lindqvist@intive.com'
        },
        {
            key: 31,
            imageUrl: '',
            imageInitials: 'AL',
            primaryText: 'Ann Lindqiest',
            secondaryText: 'UX Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'annie.lindqvist@intive.com'
        },
        {
            key: 32,
            imageUrl: '',
            imageInitials: 'AR',
            primaryText: 'Aron Reid',
            secondaryText: 'UX Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'aron.reid@intive.com'
        },
        {
            key: 33,
            imageUrl: '',
            imageInitials: 'AR',
            primaryText: 'Aaron Reed',
            secondaryText: 'UX Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'aaron.reed@intive.com'
        },
        {
            key: 34,
            imageUrl: '',
            imageInitials: 'AL',
            primaryText: 'Alix Lindberg',
            secondaryText: 'UX Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'alix.lindberg@intive.com'
        },
        {
            key: 35,
            imageUrl: '',
            imageInitials: 'AL',
            primaryText: 'Alan Lindberg',
            secondaryText: 'UX Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'alan.lindberg@intive.com'
        },
        {
            key: 36,
            imageUrl: '',
            imageInitials: 'MS',
            primaryText: 'Maor Sharit',
            secondaryText: 'UX Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'maor.sharit@intive.com'
        },
        {
            key: 37,
            imageUrl: '',
            imageInitials: 'MS',
            primaryText: 'Maorr Sherit',
            secondaryText: 'UX Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'maor.sharit@intive.com'
        },
        {
            key: 38,
            imageUrl: '',
            imageInitials: 'AL',
            primaryText: 'Alex Lindbirg',
            secondaryText: 'Software Developer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'alex.lindbirg@intive.com'
        },
        {
            key: 39,
            imageUrl: '',
            imageInitials: 'AL',
            primaryText: 'Alex Lindbarg',
            secondaryText: 'Software Developer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            email: 'alex.lindbarg@intive.com'
        }
    ];

    getPeoplesByString(text: string, maxResults: number): Promise<(IPersonaProps & { key: string | number })[]> {
        let p = this.people.filter(x => x.primaryText.indexOf(text) > -1 || x.email.indexOf(text) > -1).slice(0, maxResults);
        return new Promise<(IPersonaProps & { key: string | number })[]>((resolve, reject) => setTimeout(() => resolve(p), 2000));
    }

    getCurrentUser(success: (user: User) => void, error: (msg: string) => void): void {

        let user: User = {
            userId: 1,
            title: 'Software Developer',
            loginName: 'alex.lindbarg',
            email: 'alex.lindbarg@intive.com'
        }

        success(user);
    }

    getCurrentPersonaUser(success: (user: PersonaUser) => void, error: (msg: string) => void) {
        this.getCurrentUser(user => {
            this.getPeoplesByString(user.email, 1)
                .then(result => {

                    if (result.length === 0) {
                        error(`Cannot find user by email: ${user.email}`);
                    } else {

                        let up = result[0];

                        let personaUser: PersonaUser = {
                            userId: user.userId,
                            email: user.email,
                            title: user.title,
                            loginName: user.loginName,
                            key: up.key,
                            imageUrl: up.imageUrl,
                            imageInitials: up.imageInitials,
                            primaryText: up.primaryText,
                            secondaryText: up.title
                        };

                        success(personaUser);
                    }

                });
        }, error);
    }
}