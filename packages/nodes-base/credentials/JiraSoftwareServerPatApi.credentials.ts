import type {
	IAuthenticateGeneric,
	ICredentialTestRequest,
	ICredentialType,
	INodeProperties,
} from 'n8n-workflow';

export class JiraSoftwareServerPatApi implements ICredentialType {
	name = 'jiraSoftwareServerPatApi';

	displayName = 'Jira SW Server (PAT) API';

	documentationUrl = 'jira';

	properties: INodeProperties[] = [
		{
			displayName: 'Personal Access Token',
			name: 'personalAccessToken',
			typeOptions: { password: true },
			type: 'string',
			default: '',
		},
		{
			displayName: 'Domain',
			name: 'domain',
			type: 'string',
			default: '',
			placeholder: 'https://example.com',
		},
		{
			displayName: 'Add Proxy',
			name: 'proxySetting',
			type: 'boolean',
			default: false,
		},
		{
			displayName: 'Proxy Value',
			name: 'proxy',
			type: 'string',
			displayOptions: {
				show: {
					proxySetting: [true],
				},
			},
			default: '',
		},
	];

	authenticate: IAuthenticateGeneric = {
		type: 'generic',
		properties: {
			headers: {
				Authorization: '=Bearer {{$credentials.personalAccessToken}}',
			},
		},
	};

	test: ICredentialTestRequest = {
		request: {
			baseURL: '={{$credentials?.domain}}',
			url: '/rest/api/2/myself',
		},
	};
}
