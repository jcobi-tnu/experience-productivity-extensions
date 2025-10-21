// Copyright 2021-2023 Ellucian Company L.P. and its affiliates.

import React, { useEffect, useMemo, useState } from 'react';
import PropTypes from 'prop-types';

import { useCache, useCardInfo, useExtensionControl } from '@ellucian/experience-extension-utils';
import { Context } from '../../context-hooks/auth-context-hooks';

import { acquireToken, initializeAuthEvents, initializeMicrosoft, initializeGraphClient, login, logout } from '../util/auth';

import log from 'loglevel';
import { Client } from '@microsoft/microsoft-graph-client';
import { name, publisher } from '../../../microsoft-extension'

const logger = log.getLogger('Microsoft');

export function MicrosoftAuthProvider({ children }) {
    const { getItem: cacheGetItem, storeItem: cacheStoreItem } = useCache();
    const {
        configuration: {
            aadRedirectUrl,
            aadClientId,
            aadTenantId
        }
    } = useCardInfo();

    const [msalClient, setMsalClient] = useState();
    const [graphClient, setGraphClient] = useState();
    const [error, setError] = useState(false);
    const [loggedIn, setLoggedIn] = useState(false);
    const [state, setState] = useState('initialize');
    const [platform, setPlatform] = useState('android');
    const { setLoadingStatus } = useExtensionControl();

    useEffect(() => {

    // Check if setInvokable is available
    if (!window?.setInvokable) {
        console.error('❌ window.setInvokable is NOT available!');
        return;
    }

    // console.log('✅ window.setInvokable is available');

    // ===== CALLBACK 1: mobileLogin =====
  function mobileLogin(data) {
    

    try {
        let parsedData;

        // Parse if it's a JSON string
        if (typeof data === 'string') {
            parsedData = JSON.parse(data);
        } else {
            parsedData = data;
        }

        // console.log('🔥 Parsed data:', parsedData);

        // Extract the access token
        const accessToken = parsedData?.accessToken;

        if (!accessToken) {
            console.error('❌ No accessToken in parsed data');
            return;
        }

        // console.log('✅ Access token found');
        // console.log('   Client ID:', parsedData.clientId);
        // console.log('   Platform:', parsedData.platform);
        // console.log('   Provider:', parsedData.provider);
        // console.log('   Token length:', accessToken.length);
        // console.log('   Expires in:', parsedData.expiresIn);

        // Initialize Graph client with the token
        const options = {
            authProvider: {
                getAccessToken: () => accessToken
            }
        };

        const graphClient = Client.initWithMiddleware(options);

        if (graphClient) {
            setGraphClient(() => graphClient);
            setLoggedIn(true);
            setState('ready');
            // console.log('✅ Graph client initialized successfully!');
        } else {
            console.error('❌ Failed to initialize graph client');
        }
    } catch (error) {
        console.error('❌ Error in mobileLogin:', error);
        console.error('   Error stack:', error.stack);
        setError(true);
    }
}



    // ===== CALLBACK 2: getNewAccessToken =====
function getNewAccessToken(data) {
    // console.log('🔥🔥🔥 getNewAccessToken CALLED');
    // console.log('🔥 data type:', typeof data);
    // console.log('🔥 data raw:', data);

    try {
        const parsedData = typeof data === 'string' ? JSON.parse(data) : data;

        // console.log('🔥 Parsed data:', parsedData);

        // Capture platform if provided
        if (parsedData?.platform) {
            // console.log('📱 Platform detected:', parsedData.platform);
            setPlatform(parsedData.platform);
        }

        // ✅ Destructure accessToken to satisfy "prefer-destructuring"
        const { accessToken } = parsedData;

        if (accessToken) {
            // console.log('✅ Access token found, calling mobileLogin...');
            mobileLogin({ ...parsedData, accessToken });
        } else if (parsedData?.platform) {
            // console.log('ℹ️ Platform info only (no cached token):', parsedData.platform);
            setLoggedIn(false);
            setGraphClient(null);
            setState('ready');
            // console.log('🔄 User signed out, state updated');
        } else {
            // console.log('⚠️ Unexpected data structure:', parsedData);
        }
    } catch (error) {
        console.error('❌ Error parsing getNewAccessToken data:', error);
    }
}


    // ===== CALLBACK 3: setLoading =====
    function setLoading(status) {
        // console.log('🔥🔥🔥 setLoading CALLED with status:', status);

        if (status === 'true' || status === true) {
            setLoadingStatus(true);
        } else if (status === 'false' || status === false) {
            setState('ready');
            setLoadingStatus(false);
        }
    }

    // ===== CALLBACK 4: mobileLogOut =====
    function mobileLogOut() {
        // console.log('🔥🔥🔥 mobileLogOut CALLED');
        setState('event-logout');
    }

    // ===== REGISTERING ALL CALLBACKS =====
    // console.log('🔥 Registering mobileLogin...');
    window.setInvokable('mobileLogin', mobileLogin);

    // console.log('🔥 Registering getNewAccessToken...');
    window.setInvokable('getNewAccessToken', getNewAccessToken);

    // console.log('🔥 Registering setLoading...');
    window.setInvokable('setLoading', setLoading);

    // console.log('🔥 Registering mobileLogout...');
    window.setInvokable('mobileLogout', mobileLogOut);

    // ===== VERIFYING REGISTRATION =====
    // console.log('🔥 Verifying callbacks are accessible...');
    if (window.mobileLogin) {
        // console.log('✅ window.mobileLogin is accessible');
    } else {
        console.error('❌ window.mobileLogin is NOT accessible');
    }

    // console.log('🔥 === CALLBACK REGISTRATION COMPLETE ===');

    // catch-all logging
    const originalSetInvokable = window.setInvokable;
    window.setInvokable = function(name, callback) {
        // console.log(`🔥 setInvokable called for: ${name}`);
        return originalSetInvokable(name, (...args) => {
            // console.log(`🔥🔥🔥 Callback "${name}" invoked with args:`, args);
            return callback(...args);
        });
    };

    // ===== NOW TRYING TO GET CACHED TOKEN =====
    // console.log('🔥 Attempting to call acquireMobileToken...');
    if (window?.invokeNativeFunction) {
        const extName = `${name.replace(/ /g, '')}+${publisher.replace(/ /g, '')}`;
        // console.log('🔥 Calling userSignIn with extName:', extName);
        window.invokeNativeFunction(
            'acquireMobileToken',
            {
                randomVal: Math.random(),
                extName
            },
            false
        );
        // console.log('✅ acquireMobileToken called');

        if (window?.isInNativeApp && window.isInNativeApp()) {
            // console.log('✅ In native app, setting state to ready after 100ms');
            setTimeout(() => setState('ready'), 100);
        }
    } else {
        console.error('❌ window.invokeNativeFunction not available');
    }

    function onAuthError(error) {
    // console.log('🔥🔥🔥 onAuthError CALLED');
    console.error('❌ Auth error:', error);
    setError(true);
    setState('ready');
}

// Register AuthError
// console.log('🔥 Registering onAuthError...');
window.setInvokable('onAuthError', onAuthError);

}, []);

    // eslint-disable-next-line complexity
    useEffect(() => {
        switch (state) {
            case 'initialize':
                if (aadClientId && aadRedirectUrl && aadTenantId && !msalClient) {
                    const msalClient = initializeMicrosoft({ aadClientId, aadRedirectUrl, aadTenantId, setMsalClient });
                    setMsalClient(() => msalClient);
                    initializeAuthEvents({ setState });

                    // check if already logged in
                    (async () => {
                        if (await acquireToken({ aadClientId, aadTenantId, cacheGetItem, cacheStoreItem, msalClient, trySsoSilent: true })) {
                            setState('do-graph-initialize');
                        }
                        // check whether we are in the native app
                        else if (window?.isInNativeApp ? !window.isInNativeApp() : true) {
                            setState('ready');
                        }
                    })();
                }
                break;
            case 'do-login':
                if (aadClientId && aadRedirectUrl && aadTenantId && cacheGetItem && cacheStoreItem && msalClient) {
                    (async () => {
                        if (await login({ aadClientId, aadRedirectUrl, aadTenantId, cacheGetItem, cacheStoreItem, msalClient })) {
                            setState('do-graph-initialize');
                        } else {
                            // user likely bailed
                            setState('ready');
                        }
                    })();
                }
                break;
            case 'do-logout':
                if (aadClientId && aadRedirectUrl && aadTenantId && msalClient) {
                    (async () => {
                        await logout({ aadClientId, aadRedirectUrl, aadTenantId, msalClient });
                        setLoggedIn(false);
                        setState('ready');
                    })();
                }
                break;
            case 'do-graph-initialize':
                // check whether we are in the native app
                if (window?.isInNativeApp ? !window.isInNativeApp() : true) {
                    if (aadClientId && aadRedirectUrl && aadTenantId && msalClient) {
                        const graphClient = initializeGraphClient({ aadClientId, aadTenantId, msalClient, setError });
                        if (graphClient) {
                            setGraphClient(() => graphClient);
                            setLoggedIn(true);
                            setState('ready');
                        }
                    }
                }
                break;
            case 'event-login':
                setState('do-graph-initialize');
                break;
            case 'event-logout':
                setLoggedIn(false);
                setState('ready');
                break;
            default:
        }
    }, [aadClientId, aadRedirectUrl, aadTenantId, cacheGetItem, cacheStoreItem, msalClient, state]);

    const contextValue = useMemo(() => {
        return {
            client: graphClient,
            error,
            login: () => { setState('do-login') },
            logout: () => { setState('do-logout') },
            loggedIn,
            setLoggedIn,
            platform,
            state: state === 'ready' || state === 'do-logout' ? 'ready' : 'not-ready'
        }
    }, [graphClient, error, loggedIn, login, state, platform]);

    useEffect(() => {
        // eslint-disable-next-line no-console
        logger.debug('MicrosoftAuthProvider mounted');
        return () => {
            // eslint-disable-next-line no-console
            logger.debug('MicrosoftAuthProvider unmounted');
        }
    }, []);

    return (
        <Context.Provider value={contextValue}>
            {children}
        </Context.Provider>
    )
}

MicrosoftAuthProvider.propTypes = {
    children: PropTypes.object.isRequired
}