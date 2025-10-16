// Copyright 2021-2023 Ellucian Company L.P. and its affiliates.



module.exports = {
    name: 'Microsoft Productivity Tools',
    publisher: 'Trevecca Nazarene University',
    configuration: {
        client: [{
            key: 'aadRedirectUrl',
            label: 'Azure AD Redirect URL',
            type: 'url',
            required: true
        }, {
            key: 'aadClientId',
            label: 'Azure AD Application (Client) ID',
            type: 'text',
            required: true
        }, {
            key: 'aadTenantId',
            label: 'Azure AD Tenant ID',
            type: 'text',
            required: true
        }]
	},
    cards: [
    {
        type: 'MicrosoftToDoCard',
        source: './src/microsoft/cards/ToDo',
        title: 'Microsoft To Do',
        displayCardType: 'Microsoft ToDo',
        description: 'This card displays your tasks across lists'
    }]
}


