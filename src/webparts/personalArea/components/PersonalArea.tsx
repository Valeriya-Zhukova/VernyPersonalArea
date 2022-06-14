import * as React from 'react';
import styles from './PersonalArea.module.scss';
import { IPersonalAreaProps } from './IPersonalAreaProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IErrand, ILogin, IPersonalInfo, IVacation } from './interfaces';
import { Pivot, PivotItem } from 'office-ui-fabric-react/lib/Pivot';
import { SPFI, spfi } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/site-users/web';
import { Icon } from 'office-ui-fabric-react';
import { getSP } from '../pnpjsConfig';
import { ICamlQuery } from '@pnp/sp/lists';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from '@fluentui/react/lib/DetailsList';

export interface IAsyncAwaitPnPJsProps {
	description: string;
}

export interface IPersonalAreaState {
	personalInfo: IPersonalInfo;
	vacation: IVacation;
	errand: IErrand;
	columns: IColumn[];
}

const _styles = {
	root: [
		{
			selectors: {
				'.ms-DetailsHeader': {
					paddingTop: '0',
				},
			},
		},
	],
};

export default class PersonalArea extends React.Component<IPersonalAreaProps, IPersonalAreaState> {
	private _sp: SPFI;

	constructor(props: IPersonalAreaProps, state: IPersonalAreaState) {
		super(props);

		const _columns: IColumn[] = [
			{
				key: 'from',
				name: 'С',
				fieldName: 'from',
				minWidth: 80,
				maxWidth: 80,
			},
			{
				key: 'to',
				name: 'По',
				fieldName: 'to',
				minWidth: 80,
				maxWidth: 80,
			},
			{
				key: 'comment',
				name: 'Комментарий',
				fieldName: 'comment',
				minWidth: 220,
				maxWidth: 440,
			},
			{
				key: 'order',
				name: 'Приказ',
				fieldName: 'order',
				minWidth: 80,
				maxWidth: 80,
			},
		];

		// set initial state
		this.state = {
			personalInfo: {
				Title: 'User not found',
				post: '',
				division: '',
				Phone: '',
				Mobile: '',
				acceptd: '',
				vacation: '',
				email: '',
			},
			vacation: {
				// vacationDays: 0,
				title: '',
				list: [],
			},
			errand: {
				errandDays: 0,
				title: '',
				list: [],
			},
			columns: _columns,
		};
		this._sp = getSP();
	}

	// после монтирования компонента обновить state
	public componentDidMount(): void {
		this._getUserPersonalInfo();
		this._getUserVacationsInfo();
		this._getUserErrandsInfo();
		console.log(this.state);
	}

	public render(): React.ReactElement<IPersonalAreaProps> {
		return (
			<div className={styles.personalArea}>
				<div className={`${styles.container} ${styles.round}`}>
					<div className={`${styles.profileWrapper}`}>
						<header className={`${styles.topHeader}`}>
							<h3 className={`${styles.header}`}>Личный кабинет</h3>
						</header>
						<Pivot>
							<PivotItem headerText='Инфо'>
								<div className={`${styles.tabContent}`}>
									<div className={`${styles.tabContentWrapper}`}>
										<div className={`${styles.left}`}>
											<section className={`${styles.generalInfo} ${styles.round}`}>
												<div className={`${styles.card}`}>
													<div className={`${styles.cardHeader}`}>
														<div className={`${styles.name}`}>
															<span>{this.state.personalInfo.Title}</span>
														</div>
													</div>
													<div className={`${styles.cardBody}`}>
														<div className={`${styles.flexRow}`}>
															<div className={`${styles.image}`}>
																<img
																	className={`${styles.photo}`}
																	// src={`/_vti_bin/DelveApi.ashx/people/profileimage?size=L&amp;userId=${this.state.login.userName}`}
																	src={`/_layouts/15/userphoto.aspx?size=L&username=${this.props.login}`}
																/>
															</div>
														</div>
													</div>
												</div>
											</section>

											<section className={`${styles.details} ${styles.round}`}>
												<div className={`${styles.card}`}>
													<div className={`${styles.cardHeader}`}>
														<i className="${getIconclassName('Info')} ${styles.headerIcon}"></i>Личная информация
													</div>
													<div className={`${styles.cardBody}`}>
														<ul className={`${styles.detailsList}`}>
															<li id='job'>
																<div className={`${styles.detailsItem}`}>
																	<div className={`${styles.detailsItemTitle}`}>Должность</div>
																	<div className={`${styles.detailsItemValue}`}>{this.state.personalInfo.post}</div>
																</div>
															</li>
															<li id='department'>
																<div className={`${styles.detailsItem}`}>
																	<div className={`${styles.detailsItemTitle}`}>Департамент</div>
																	<div className={`${styles.detailsItemValue}`}>{this.state.personalInfo.division}</div>
																</div>
															</li>
															<li id='internalNumber'>
																<div className={`${styles.detailsItem}`}>
																	<div className={`${styles.detailsItemTitle}`}>Внутренний номер</div>
																	<div className={`${styles.detailsItemValue}`}>{this.state.personalInfo.Phone}</div>
																</div>
															</li>
															<li id='phoneNumber'>
																<div className={`${styles.detailsItem}`}>
																	<div className={`${styles.detailsItemTitle}`}>Номер телефона</div>
																	<div className={`${styles.detailsItemValue}`}>{this.state.personalInfo.Mobile}</div>
																</div>
															</li>
															<li id='email'>
																<div className={`${styles.detailsItem}`}>
																	<div className={`${styles.detailsItemTitle}`}>Email</div>
																	<div className={`${styles.detailsItem}value`}>{this.state.personalInfo.email}</div>
																</div>
															</li>
															<li id='worksFrom'>
																<div className={`${styles.detailsItem}`}>
																	<div className={`${styles.detailsItemTitle}`}>Работает в ТОО Verny Capital c</div>
																	<div className={`${styles.detailsItemValue}`}>{this.state.personalInfo.acceptd}</div>
																</div>
															</li>
															<li id='currentJobFrom'>
																<div className={`${styles.detailsItem}`}>
																	<div className={`${styles.detailsItemTitle}`}>Дней отпуска</div>
																	<div className={`${styles.detailsItemValue}`}>{this.state.personalInfo.vacation}</div>
																</div>
															</li>
														</ul>
													</div>
												</div>
											</section>
										</div>

										<div className={`${styles.right}`}>
											<section className={`${styles.links}`}>
												<div className={`${styles.card} ${styles.cardBlue}`}>
													<div className={`${styles.cardHeader}`}>
														<i className="${getIconClassName('Link')} ${styles.headerIcon}"></i>
														Полезные ссылки
													</div>
													<div className={`${styles.cardBody}`}>
														<a
															href='https://vernycapital.sharepoint.com/howto/Shared%20Documents/Forms/AllItems.aspx?viewid=a8fac01f%2D5f1b%2D4789%2Db2e9%2Dd6daf52fbbed'
															target='_blank'
															className={`${styles.link}`}
														>
															<button type='button' className={`${styles.btn} ${styles.btnDark}`}>
																<Icon iconName='PageList' className={`${styles.btnIcon}`}></Icon>
																Должностная Инструкция
															</button>
														</a>
														<a
															href='https://vernycapital.sharepoint.com/howto/Shared%20Documents/Forms/AllItems.aspx?viewid=a8fac01f%2D5f1b%2D4789%2Db2e9%2Dd6daf52fbbed'
															target='_blank'
															className={`${styles.link}`}
														>
															<button type='button' className={`${styles.btn} ${styles.btnDark}`}>
																<Icon iconName='ChangeEntitlements' className={`${styles.btnIcon}`}></Icon>
																Центр Инструкций
															</button>
														</a>
													</div>
												</div>
											</section>

											<section className={`${styles.hr}`}>
												<div className={`${styles.card} ${styles.cardBlue}`}>
													<div className={`${styles.cardHeader}`}>
														<Icon iconName='Processing' className={`${styles.btnIcon}`}></Icon>
														Системное управление сервисами
													</div>
													<div className={`${styles.cardBody}`}>
														<a
															href='https://bpm.vernycapital.com/itsm/itsm_form?uid=11e9ce85-ad35-47e0-91e8-3750fb1f6296'
															target='_blank'
															className={`${styles.link}`}
														>
															<button type='button' className={`${styles.btn} ${styles.btnDark}`}>
																<Icon iconName='PaymentCard' className={`${styles.btnIcon}`}></Icon>
																<i className="${getIconClassName('PaymentCard')} ${styles.btnIcon}"></i>
																Заявка на оплату
															</button>
														</a>
														<a href='https://bpm.vernycapital.com' target='_blank' className={`${styles.link}`}>
															<button type='button' className={`${styles.btn} ${styles.btnDark}`}>
																<Icon iconName='DocumentReply' className={`${styles.btnIcon}`}></Icon>
																Согласование документов
															</button>
														</a>
													</div>
												</div>
											</section>

											<section className={`${styles.hr}`}>
												<div className={`${styles.card} ${styles.cardBlue}`}>
													<div className={`${styles.cardHeader}`}>
														<Icon iconName='People' className={`${styles.btnIcon}`}></Icon>
														HR
													</div>
													<div className={`${styles.cardBody}`}>
														<a href='https://bpm.vernycapital.com' target='_blank' className={`${styles.link}`}>
															<button type='button' className={`${styles.btn} ${styles.btnDark}`}>
																<Icon iconName='ClipboardList' className={`${styles.btnIcon}`}></Icon>
																PAS
															</button>
														</a>
													</div>
												</div>
											</section>
										</div>
									</div>
								</div>
							</PivotItem>
							<PivotItem headerText='Отпуска'>
								<div className={`${styles.tabContent}`}>
									<div className={`${styles.tabContentWrapper}`}>
										<div className={`${styles.left}`}>
											<section className={`${styles.generalInfo} ${styles.round}`}>
												<div className={`${styles.card}`}>
													<div className={`${styles.cardHeader}`}>
														<div className={`${styles.name}`}>
															<span>Дней отпуска: {this.state.personalInfo.vacation}</span>
														</div>
													</div>
													<div className={`${styles.cardBody}`}>
														<DetailsList
															items={this.state.vacation.list}
															columns={this.state.columns}
															layoutMode={DetailsListLayoutMode.justified}
															isHeaderVisible={true}
															selectionMode={SelectionMode.none}
															compact={true}
															styles={_styles}
														/>
													</div>
												</div>
											</section>
										</div>
										<div className={`${styles.right}`}>
											<section className={`${styles.links} ${styles.round}`}>
												<div className={`${styles.card} ${styles.cardBlue}`}>
													<div className={`${styles.cardHeader}`}>
														<div className={`${styles.name}`}>
															<span>Полезные ссылки</span>
														</div>
													</div>
													<div className={`${styles.cardBody}`}>
														<a
															href={`https://vernycapital.sharepoint.com/Lists/OutOf/AllItems.aspx?FilterField1=login&FilterValue1=${this.props.login}&useFiltersInViewXml=1&FilterField2=Type&FilterValue2=%D0%9E%D1%82%D0%BF%D1%83%D1%81%D0%BA&FilterType2=Choice&FilterOp2=In`}
															target='_blank'
															className={`${styles.link}`}
														>
															<button type='button' className={`${styles.btn} ${styles.btnLightGreen}`}>
																<Icon iconName='Vacation' className={`${styles.btnIcon}`}></Icon>
																Посмотреть список отпусков
															</button>
														</a>
														<a
															href='https://bpm.vernycapital.com/itsm/itsm_main#catalogue'
															target='_blank'
															className={`${styles.link}`}
														>
															<button type='button' className={`${styles.btn} ${styles.btnDark}`}>
																<Icon iconName='ActivateOrders' className={`${styles.btnIcon}`}></Icon>
																Подать заявку на отпуск
															</button>
														</a>
													</div>
												</div>
											</section>
										</div>
									</div>
								</div>
							</PivotItem>
							<PivotItem headerText='Командировки'>
								<div className={`${styles.tabContent}`}>
									<div className={`${styles.tabContentWrapper}`}>
										<div className={`${styles.left}`}>
											<section className={`${styles.generalInfo} ${styles.round}`}>
												<div className={`${styles.card}`}>
													<div className={`${styles.cardHeader}`}>
														<div className={`${styles.name}`}>
															<span>Командировок с начала года: {this.state.errand.errandDays}</span>
														</div>
													</div>
													<div className={`${styles.cardBody}`}>
														<DetailsList
															items={this.state.errand.list}
															columns={this.state.columns}
															layoutMode={DetailsListLayoutMode.justified}
															isHeaderVisible={true}
															selectionMode={SelectionMode.none}
															compact={true}
															styles={_styles}
														/>
													</div>
												</div>
											</section>
										</div>

										<div className={`${styles.right}`}>
											<section className={`${styles.links} ${styles.round}`}>
												<div className={`${styles.card} ${styles.cardBlue}`}>
													<div className={`${styles.cardHeader}`}>
														<div className={`${styles.name} `}>
															<span>Полезные ссылки</span>
														</div>
													</div>
													<div className={`${styles.cardBody}`}>
														<a
															href={`https://vernycapital.sharepoint.com/Lists/OutOf/AllItems.aspx?FilterField1=login&FilterValue1=${this.props.login}&useFiltersInViewXml=1&FilterField2=Type&FilterValue2=%D0%9A%D0%BE%D0%BC%D0%B0%D0%BD%D0%B4%D0%B8%D1%80%D0%BE%D0%B2%D0%BA%D0%B0&FilterType2=Choice&FilterOp2=In`}
															target='_blank'
															className={`${styles.link}`}
														>
															<button type='button' className={`${styles.btn} ${styles.btnLightBlue}`}>
																<Icon iconName='Arrivals' className={`${styles.btnIcon}`}></Icon>
																Посмотреть список командировок
															</button>
														</a>
														<a
															href='https://bpm.vernycapital.com/itsm/itsm_form?uid=5a1d9e58-11b5-43b4-ad9f-f1257505b179'
															target='_blank'
															className={`${styles.link}`}
														>
															<button type='button' className={`${styles.btn} ${styles.btnDark}`}>
																<Icon iconName='ActivateOrders' className={`${styles.btnIcon}`}></Icon>
																Подать заявку на командировку
															</button>
														</a>
													</div>
												</div>
											</section>
										</div>
									</div>
								</div>
							</PivotItem>
						</Pivot>
					</div>
				</div>
			</div>
		);
	}

	private _getUserPersonalInfo = async (): Promise<void> => {
		try {
			const sp = spfi(this._sp);

			const _email = (await sp.web.currentUser()).Email;

			const item: object = await sp.web.lists
				.getByTitle('profile')
				.items.select('Title', 'post', 'Phone', 'Mobile', 'division', 'acceptd', 'vacation')
				.filter(`login eq \'${this.props.login}\'`)();

			console.log(item);

			const _personalInfo = {
				Title: item[0]?.Title,
				post: item[0]?.post,
				Phone: item[0]?.Phone,
				Mobile: item[0]?.Mobile,
				division: item[0]?.division,
				acceptd: item[0]?.acceptd,
				vacation: item[0]?.vacation,
				email: _email,
			};

			this.setState({ personalInfo: _personalInfo });
			console.log(this.state);
			console.log(this.props);
		} catch (err) {
			console.log(err);
		}
	};

	// отпуска
	private _getUserVacationsInfo = async (): Promise<void> => {
		try {
			const sp = spfi(this._sp);

			// TODO: CamlQuery для получения списка отпусков пользователя
			const camlVacations: ICamlQuery = {
				ViewXml: `<View>
				<Query>
					<Where>
						<And>
							<Eq>
								<FieldRef Name='Title' />
								<Value Type='Text'>asmyshlyayev@vernycapital.com</Value>
							</Eq>
							<And>
								<Gt>
									<FieldRef Name='From' />
									<Value IncludeTimeValue='TRUE' Type='DateTime'>2022-01-01T13:40:32Z</Value>
								</Gt>
								<Eq>
									<FieldRef Name='Type' />
									<Value Type='Choice'>Отпуск</Value>
								</Eq>
							</And>
						</And>
					</Where>
				</Query>
			</View>`,
			};

			const itemsCamlVacations: any[] = await sp.web.lists.getByTitle('Нет в офисе').getItemsByCAMLQuery(camlVacations);

			console.log(camlVacations);

			// создает массив отпусков для вывода
			const vacationList = itemsCamlVacations.map((item) => {
				const _from = new Date(item.From);
				const _to = new Date(item.To);
				const currentObj = {
					from: `${_from.getDate()}.${_from.getMonth() + 1}.${_from.getFullYear()}`,
					to: `${_to.getDate()}.${_to.getMonth() + 1}.${_to.getFullYear()}`,
					comment: item.Commnet,
					order: item.OrderID,
				};

				return currentObj;
			});

			// обновление state
			const _vacation = {
				// vacationDays: _vacationDays,
				title: itemsCamlVacations[0]?.Title,
				list: vacationList,
			};

			console.log(_vacation);

			this.setState({ vacation: _vacation });
		} catch (error) {
			console.log(error);
		}
	};

	// командировки
	private _getUserErrandsInfo = async (): Promise<void> => {
		try {
			const sp = spfi(this._sp);

			// TODO: CamlQuery для получения списка командировок пользователя
			const camlErrands: ICamlQuery = {
				ViewXml: `<View>
				<Query>
					<Where>
						<And>
							<Eq>
								<FieldRef Name='login' />
								<Value type='text'>asmyshlyayev@vernycapital.com</Value>
							</Eq>
							<And>
								<Gt>
									<FieldRef Name='From' />
									<Value IncludeTimeValue='TRUE' Type='DateTime'>2022-01-01T13:40:32Z</Value>
								</Gt>
								<Eq>
									<FieldRef Name='Type' />
									<Value Type='Choice'>Отпуск</Value>
								</Eq>
							</And>
						</And>
					</Where>
				</Query>
			</View>`,
			};

			const itemsCamlErrands: any[] = await sp.web.lists.getByTitle('Нет в офисе').getItemsByCAMLQuery(camlErrands);

			console.log(itemsCamlErrands);

			const _errandDays = itemsCamlErrands.length;

			const vacationList = itemsCamlErrands.map((item) => {
				const _from = new Date(item.From);
				const _to = new Date(item.To);
				const currentObj = {
					from: `${_from.getDate()}.${_from.getMonth() + 1}.${_from.getFullYear()}`,
					to: `${_to.getDate()}.${_to.getMonth() + 1}.${_to.getFullYear()}`,
					comment: item.Commnet,
					order: item.OrderID,
				};

				return currentObj;
			});

			const _errand = {
				errandDays: _errandDays,
				title: itemsCamlErrands[0]?.Title,
				list: vacationList,
			};

			console.log(_errand);

			this.setState({ errand: _errand });
		} catch (error) {
			console.log(error);
		}
	};
}
