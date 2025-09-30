import React, { useEffect, useMemo, useState } from 'react';
import { useIntl as useCardIntl } from '../context-hooks/card-context-hooks';
import { useAuth } from '../context-hooks/auth-context-hooks';
import { useExtensionControl, useUserInfo } from '@ellucian/experience-extension-utils';
import PropTypes from 'prop-types';
import { withStyles } from '@ellucian/react-design-system/core/styles';
import {
    colorBrandNeutral250,
    colorBrandNeutral300,
    fontWeightNormal,
    spacing30,
    spacing40
} from '@ellucian/react-design-system/core/styles/tokens';
import {
    Checkbox,
    Illustration,
    IMAGES,
    Typography,
    Divider,
    List,
    ListItem,
    ListItemText,
    IconButton
} from '@ellucian/react-design-system/core';
import { Icon } from '@ellucian/ds-icons/lib';
import SignInButton from './SignInButton';
import SignOutButton from './SignOutButton';
import { useToDo } from '../context-hooks/todo-context-hooks';
import { log } from 'loglevel';

const customId = 'BasicIconButtons';

const styles = () => ({
    card: {
        flex: '1 0 auto',
        width: '100%',
        display: 'flex',
        padding: `0 ${spacing40} ${spacing40} ${spacing40}`,
        flexFlow: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        '& > *': {
            marginBottom: spacing40
        }
    },
    topBar: {
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        width: '100%'
    },
    content: {
        display: 'flex',
        flexDirection: 'column',
        marginLeft: spacing40,
        marginRight: spacing40,
        '& :first-child': {
            paddingTop: '0px'
        },
        '& hr:last-of-type': {
            display: 'none'
        }
    },
    row: {
        display: 'flex',
        alignItems: 'center',
        paddingTop: spacing30,
        paddingBottom: spacing30,
        paddingLeft: spacing30,
        paddingRight: spacing30,
        '&:hover': {
            backgroundColor: colorBrandNeutral250
        }
    },
    noWrap: {
        overflow: 'hidden',
        whiteSpace: 'nowrap',
        textOverflow: 'ellipsis'
    },
    fontWeightNormal: {
        fontWeight: fontWeightNormal
    },
    divider: {
        marginTop: '0px',
        marginBottom: '0px',
        backgroundColor: colorBrandNeutral300
    },
    logoutBox: {
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        marginTop: spacing40,
        marginBottom: spacing40
    },
    settings: {
        marginTop: spacing40
    },
    refreshButton: {
        margin: spacing40
    }
});

function ToDo({ classes }) {
    const { intl, components } = useCardIntl();
    const { setErrorMessage, setLoadingStatus } = useExtensionControl();
    const { locale } = useUserInfo();
    const { error: authError, login, loggedIn, logout, state: authState } = useAuth();
    const { tasks, state, error: todoError, refresh, toggleComplete } = useToDo();

    const [showCompleted, setShowCompleted] = useState(false);
    const [displayState, setDisplayState] = useState('loading');

    const isLoading = state === 'load' || state === 'refresh';

    const formatter = useMemo(
        () => new Intl.DateTimeFormat(locale, { dateStyle: 'medium' }),
        [locale]
    );

    const visible = useMemo(() => {
        return showCompleted ? tasks : tasks.filter(t => t.status !== 'completed');
    }, [tasks, showCompleted]);

    useEffect(() => {
        if (authError || todoError) {
            setErrorMessage({
                headerMessage: intl.formatMessage({ id: 'error.contentNotAvailable' }),
                textMessage: intl.formatMessage({ id: 'error.contactYourAdministrator' }),
                iconName: 'warning'
            });
        } else if (loggedIn === false && authState === 'ready') {
            setDisplayState('loggedOut');
        } else if (state === 'load') {
            setDisplayState('loading');
        } else if ((state === 'loaded' || state === 'refresh') && loggedIn) {
            setDisplayState('loaded');
        } else if (state && state.error) {
            setDisplayState('error');
        }
    }, [authError, todoError, loggedIn, authState, state, intl, setErrorMessage]);

    useEffect(() => {
        setLoadingStatus(displayState === 'loading');
    }, [displayState, setLoadingStatus]);

    if (displayState === 'loggedOut') {
        return (
            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', padding: 16 }}>
                <Illustration name={IMAGES.ID_BADGE} />
                <Typography variant="h3" component="div" style={{ margin: '12px 0' }}>
                    {intl.formatMessage({ id: 'google.permissionsRequested' })}
                </Typography>
                <SignInButton onClick={login} />
            </div>
        );
    }

    if ((authError || todoError) && !isLoading) {
        return (
            <div>
                <Typography variant="body2">
                    {intl.formatMessage({ id: 'common.error' }, { defaultMessage: 'Something went wrong' })}
                </Typography>
                <div style={{ marginTop: 8 }}>
                    <IconButton
                        color="gray"
                        id={`${customId}_RetryButton`}
                        onClick={refresh}
                        aria-label={intl.formatMessage({ id: 'common.retry' }, { defaultMessage: 'Retry' })}
                        disabled={isLoading}
                        className={classes.refreshButton}
                    >
                        <Icon name="refresh" />
                    </IconButton>
                </div>
            </div>
        );
    }

    if (displayState === 'loading') {
        return (
            <Typography variant="body2">
                {intl.formatMessage({ id: 'common.loading' }, { defaultMessage: 'Loading…' })}
            </Typography>
        );
    }

    const emptyMsg = showCompleted ? components?.noTasks : components?.noActiveTasks;

    if (displayState === 'loaded' && visible.length === 0) {
        return (
            <div className={classes.card}>
                <div className={classes.topBar}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                        <Checkbox checked={showCompleted} onChange={e => setShowCompleted(e.target?.checked ?? false)} />
                        <Typography variant="body2">
                            {intl.formatMessage({ id: 'microsoft.todo.showCompleted' }, { defaultMessage: 'Show completed' })}
                        </Typography>
                    </div>
                    <IconButton
                        color="gray"
                        id={`${customId}_PrimaryButton`}
                        onClick={refresh}
                        aria-label="Refresh List"
                        disabled={isLoading}
                        className={classes.refreshButton}
                    >
                        <Icon name="refresh" />
                    </IconButton>
                </div>

                <Typography variant="h4" component="div" style={{ marginBottom: 4 }}>
                    {intl.formatMessage({ id: emptyMsg?.titleId || 'microsoft.noTasksTitle' })}
                </Typography>
                <Typography variant="body2" style={{ opacity: 0.75, marginBottom: 12 }}>
                    {intl.formatMessage({ id: emptyMsg?.messageId || 'microsoft.noTasksMessage' })}
                </Typography>
                <Divider style={{ margin: '16px 0' }} />
                <div className={classes.logoutBox}>
                    <SignOutButton onClick={logout} />
                </div>
            </div>
        );
    }

    return (
        <div className={classes.card}>
            <div className={classes.topBar}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                    <Checkbox checked={showCompleted} onChange={e => setShowCompleted(e.target?.checked ?? false)} />
                    <Typography variant="body2">
                        {intl.formatMessage({ id: 'microsoft.todo.showCompleted' }, { defaultMessage: 'Show completed' })}
                    </Typography>
                </div>
                <IconButton
                    color="gray"
                    id={`${customId}_PrimaryButton`}
                    onClick={refresh}
                    aria-label="Refresh List"
                    disabled={isLoading}
                    className={classes.refreshButton}
                >
                    <Icon name="refresh" />
                </IconButton>
            </div>

            <List dense>
                {visible.map(t => {
                    const labelId = `todo-item-${t.id}`;
                    const primary = t.link ? (
                        <a href={t.link} target="_blank" rel="noopener noreferrer">{t.title}</a>
                    ) : t.title;

                    const secondary = t.due
                        ? `${intl.formatMessage({ id: 'microsoft.todo.due', defaultMessage: 'Due' })}: ${formatter.format(new Date(t.due))}${t.listName ? ` • ${t.listName}` : ''}`
                        : `${intl.formatMessage({ id: 'microsoft.todo.noDue', defaultMessage: 'No due date' })}${t.listName ? ` • ${t.listName}` : ''}`;

                    return (
                        <ListItem key={t.id} divider>
                            <ListItemText id={labelId} primary={primary} secondary={secondary} />
                            <Checkbox
                                checked={t.status === 'completed'}
                                onChange={() => toggleComplete(t)}
                                inputProps={{ 'aria-labelledby': labelId }}
                            />
                        </ListItem>
                    );
                })}
            </List>

            <Divider style={{ margin: '12px 0' }} />
            <SignOutButton onClick={logout} />
        </div>
    );
}

ToDo.propTypes = {
    classes: PropTypes.object.isRequired
};

export default withStyles(styles)(ToDo);
