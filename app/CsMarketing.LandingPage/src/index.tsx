import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { LandingPage } from './components/landingpage/LandingPage';

const landingPageId = 'marketing-landing-page';

function start(): void {
    let landingPage = document.getElementById(landingPageId);
    if (landingPage) {
        ReactDOM.render(
            <Fabric>
                <LandingPage />
            </Fabric>,
            landingPage);
    }
}

// Start the application.
start();
