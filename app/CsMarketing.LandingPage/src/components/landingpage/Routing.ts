export const landingPage = 'app/landing-page';
export const adHocReportFormNew = 'app/forms/ad-hoc-report';


export function isLandingPage(): boolean {
    return !parent.location.hash || parent.location.hash === '' || parent.location.hash.indexOf(landingPage) > -1;
}

export function isAdHocReportFormNew(): boolean {
    return parent.location.hash.indexOf(adHocReportFormNew) > -1;
}