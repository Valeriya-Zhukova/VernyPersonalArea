import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PersonalAreaWebPartStrings';
import PersonalArea from './components/PersonalArea';
import { IPersonalAreaProps } from './components/IPersonalAreaProps';

import { getSP } from './pnpjsConfig';

export default class PersonalAreaWebPart extends BaseClientSideWebPart<IPersonalAreaProps> {
	public async onInit(): Promise<void> {
		await super.onInit();
		//Initialize our _sp object that we can then use in other packages without having to pass around the context.
		getSP(this.context);
	}

	public render(): void {
		const element: React.ReactElement<IPersonalAreaProps> = React.createElement(PersonalArea, {
			description: this.properties.description,
			login: this.context.pageContext.user.loginName,
		});

		ReactDom.render(element, this.domElement);
	}

	protected onDispose(): void {
		ReactDom.unmountComponentAtNode(this.domElement);
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0');
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		return {
			pages: [
				{
					header: {
						description: strings.PropertyPaneDescription,
					},
					groups: [
						{
							groupName: strings.BasicGroupName,
							groupFields: [
								PropertyPaneTextField('description', {
									label: strings.DescriptionFieldLabel,
								}),
							],
						},
					],
				},
			],
		};
	}
}
