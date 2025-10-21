import React from 'react';
import PropTypes from 'prop-types';
import { Button } from '@ellucian/react-design-system/core';
import { withStyles } from '@ellucian/react-design-system/core/styles';
import { spacing30 } from '@ellucian/react-design-system/core/styles/tokens';
import { useComponents, useIntl } from '../context-hooks/card-context-hooks';
import { useCardInfo } from '@ellucian/experience-extension-utils';
import { microsoftScopes } from '../microsoft/util/auth';
import { name, publisher } from '../../microsoft-extension';
import { useAuth } from '../context-hooks/auth-context-hooks';

const styles = () => ({
    button: {
        cursor: 'pointer'
    },
    image: {
        marginRight: spacing30
    }
});

function SignInButton({ classes, onClick }) {
    const { intl } = useIntl();
    const { buttonImage } = useComponents();
    const { configuration: { aadClientId, aadTenantId } } = useCardInfo();
    const { platform } = useAuth();

    const handleClick = () => {
        if (window?.isInNativeApp && window.isInNativeApp()) {
            const scopeString = microsoftScopes.join(' ');

            const params = {
                clientId: aadClientId,
                // eslint-disable-next-line camelcase
                response_type: 'token',
                scope: scopeString,
                provider: 'microsoft',
                authority: `https://login.microsoftonline.com/${aadTenantId}/oauth2/v2.0/authorize`,
                tokenUrl: `https://login.microsoftonline.com/${aadTenantId}/oauth2/v2.0/token`,
                platform: platform,
                extName: `${name.replace(/ /g, '')}+${publisher.replace(/ /g, '')}`
            };

            // console.log('🔥 ToDo card calling userSignIn:', JSON.stringify(params, null, 2));

            window.invokeNativeFunction('userSignIn', params, true);
        } else {
            onClick();
        }
    };

    return (
        <Button className={classes.button} color="secondary" onClick={handleClick}>
            <img className={classes.image} src={buttonImage} alt="buttonImage" />
            {intl.formatMessage({ id: 'signIn' })}
        </Button>
    );
}

SignInButton.propTypes = {
    classes: PropTypes.object.isRequired,
    onClick: PropTypes.func.isRequired
};

export default withStyles(styles)(SignInButton);