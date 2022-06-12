export interface IPersonalInfo {
	Title: string | string | '';
	post: string | '';
	division: string | '';
	Phone: string | '';
	Mobile: string | '';
	acceptd: string | '';
	vacation: string | '';
	email: string | '';
}

export interface ILogin {
	userName: string | '';
	email: string | '';
}

export interface IVacation {
	// vacationDays: number | null;
	title: string | null;
	list: any[] | null;
}

export interface IErrand {
	errandDays: number | null;
	title: string | null;
	list: any[] | null;
}
