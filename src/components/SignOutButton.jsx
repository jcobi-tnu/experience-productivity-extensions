import React from 'react';
import PropTypes from 'prop-types';
import classnames from 'classnames';
import { Button } from '@ellucian/react-design-system/core';
import { withStyles } from '@ellucian/react-design-system/core/styles';
import { spacing30 } from '@ellucian/react-design-system/core/styles/tokens';
import { useComponents, useIntl } from '../context-hooks/card-context-hooks';
import { name, publisher } from '../../microsoft-extension';

const styles = () => ({
    button: {
        cursor: 'pointer'
    },
    image: {
        marginRight: spacing30
    }
});

function SignOutButton({ classes, className = '', onClick }) {
    const { intl } = useIntl();
    const { buttonImage } = useComponents();

    const handleClick = () => {
        if (window?.isInNativeApp && window.isInNativeApp()) {
            const params = {
                extName: `${name.replace(/ /g, '')}+${publisher.replace(/ /g, '')}`
            };

            // console.log('🔥 ToDo card calling userSignOut:', JSON.stringify(params, null, 2));
            // ← Flag is false
            window.invokeNativeFunction('userSignOut', params, false);
        } else {
            onClick();
        }
    };

    return (
        <Button
            className={classnames(className, classes.button)}
            color="secondary"
            onClick={handleClick}
        >
            <img className={classes.image} src={buttonImage} alt="buttonImage" />
            {intl.formatMessage({ id: 'signOut' })}
        </Button>
    );
}

SignOutButton.propTypes = {
    classes: PropTypes.object.isRequired,
    className: PropTypes.string,
    onClick: PropTypes.func.isRequired
};

export default withStyles(styles)(SignOutButton);