import { type } from '@testing-library/user-event/dist/type';
import React from 'react';

var XLSX: any;
if (typeof require !== 'undefined')
	XLSX = require('xlsx');

type allGroups = {
	[key: string]: stydyWeek
}

type stydyWeek = {
	_debagProps: {
		_referenceKey: string | undefined;
		_referenceKey_isСorrect: string | undefined;
	}
	days: [
		monday: stydyDay | undefined,
		tuesday: stydyDay | undefined,
		wednesday: stydyDay | undefined,
		thursday: stydyDay | undefined,
		friday: stydyDay | undefined,
		saturday: stydyDay | undefined
	]
}

type stydyDay = {
	lessons: [
		first: lesson | undefined,
		second: lesson | undefined,
		third: lesson | undefined,
		fourth: lesson | undefined,
		fifth: lesson | undefined
	]
}

type lesson = {
	_debagProps: {
		_fullСontentСell: string | undefined;
	}
	data: {
		lessonName: string
		teacherInfo: string
		cabinet: string
		time: string
	} | undefined
}

type cell = {
	v: string
	w: string
	t: string
}

type merge =
	{
		s: { r: {}, c: {} },
		e: { r: {}, c: {} }
	}

type list = {
	'!merges': any;
	[key: string]: cell
}

var alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ'];

var pattern = /^[А-Я]+$/;
var patternEng = /^[A-Z]+$/i;
var patternNum = /^[0-9]+$/;

var workSheet: list;

var allGroups: allGroups = {};

function App() {

	const checkingMerged = (str: string, num: number) => {
		let merges: merge[] = workSheet['!merges'];
		let isMergeCell: boolean = false;
		// num = -1
		merges.map(el => {
			if ((el.s.c === alphabet.indexOf(str) && el.e.c === alphabet.indexOf(str)) && (el.s.r === num - 1 && el.e.r === num + 1 - 1) ||
				(el.s.c === alphabet.indexOf(str) && el.e.c === alphabet.indexOf(str)) && (el.s.r === num - 1 && el.e.r === num + 9 - 1)) {
				isMergeCell = true
			}
		})
		return isMergeCell
	}

	const defineGroups = () => {
		for (let curKey in workSheet) {
			let letter: number = 0, num: number = 0;
			if (workSheet[curKey].w && workSheet[curKey].w.length < 10) {
				for (let i = 0; i <= workSheet[curKey].w.length; i++) {
					if (pattern.test(workSheet[curKey].w[i])) {
						letter++;
					}
					else if (patternNum.test(workSheet[curKey].w[i]))
						num++;
				}
			}
			if (letter != 0 && num != 0 && letter < 5 && num < 5) {
				// allGroups[workSheet[curKey].w]._debagProps._referenceKey = curKey;
				let referСell_number: string = '', referСell_letter: string = '';
				for (let i = 0; i <= curKey.length; i++) {
					if (patternEng.test(curKey[i])) {
						referСell_letter += curKey[i]
					}
					else if (patternNum.test(curKey[i])) {
						referСell_number += curKey[i]
					}
					if (curKey.length - i === 1)
						break
				}
				allGroups[workSheet[curKey].w] = defineSchedule(referСell_letter, referСell_number, workSheet[curKey].w);
			}
		}


	}

	const defineSchedule = (referСell_letter: string, referСell_number: string, group: string) => {
		var stydyWeek: stydyWeek = {
			_debagProps: { _referenceKey: undefined, _referenceKey_isСorrect: undefined },
			days: [undefined, undefined, undefined, undefined, undefined, undefined]
		};
		// var grops: allGroups = {};
		let iCurrentWeekDay = 0;
		for (let iDay = +referСell_number + 1; iDay <= +referСell_number + 60; iDay += 10) {
			stydyWeek.days[iCurrentWeekDay] = defineDay(iDay, referСell_letter, referСell_number, group, iCurrentWeekDay)
			iCurrentWeekDay++;
		}

		return stydyWeek;
	}

	const defineDay = (referDay: number, referСell_letter: string, referСell_number: string, group: string, weekDay: number) => {
		var stydyDay: stydyDay = {
			lessons: [undefined, undefined, undefined, undefined, undefined,]
		};
		let iCurrentDay = 0;
		for (let iLesson = referDay; iLesson < referDay + 10; iLesson += 2) {
			let currentValue: string | undefined = workSheet[referСell_letter + iLesson] ? workSheet[referСell_letter + iLesson].w : undefined;
			let nextCell: number = alphabet.indexOf(referСell_letter) + 1;

			let incorrectData = (currentValue && currentValue.length < 6);
			let emptyData = (!currentValue && workSheet[alphabet[nextCell] + iLesson] && workSheet[alphabet[nextCell] + iLesson].w.length > 6);

			if (incorrectData || emptyData) {
				let newStr = alphabet.indexOf(referСell_letter);
				referСell_letter = alphabet[newStr + 1];
				iLesson = referDay - 2;
				// 	// if (isMerged(startStrVal, i, workSheet['!merges']))
				// 	// 	console.log('ЯЧЕЙКА ЕДИНА')
				// 	// !!! error
			} else {
				// allGroups[group]._debagProps._referenceKey_isСorrect = referСell_letter;

				stydyDay.lessons[iCurrentDay] = defineLesson(iLesson, referСell_letter, group, weekDay, currentValue)
				iCurrentDay++;
			}
		}
		return stydyDay
	}

	const defineLesson = (referLesson: number, referСell_letter: string, group: string, weekDay: number, data: string | undefined) => {
		let lesson: lesson = { _debagProps: { _fullСontentСell: '' }, data: undefined };
		// referLesson += 1;

		if (checkingMerged(referСell_letter, referLesson)) {
			if (data) {
				lesson._debagProps._fullСontentСell = data + 'merge';
			} else {
				lesson._debagProps._fullСontentСell = 'void merge';
			}
		} else {
			for (let iCell = referLesson; iCell <= referLesson + 1; iCell++) {
				let currentValue: string | undefined = workSheet[referСell_letter + iCell] ? workSheet[referСell_letter + iCell].w : undefined;
				if (data && currentValue) {
					lesson._debagProps._fullСontentСell += currentValue
				} else {
					lesson._debagProps._fullСontentСell += 'void';
				}
			}
		}
		return lesson;
	}

	const handleChange = async (e: React.ChangeEvent<HTMLInputElement>) => {

		if (!e.target.files) {
			return;
		}
		const file = e.target.files[0];
		const data = await file.arrayBuffer();
		const workBook = XLSX.read(data, {
			raw: true
		});

		workSheet = workBook.Sheets[workBook.SheetNames[0]]

		delete workSheet['!margins'];
		// delete workSheet['!merges'];
		delete workSheet['!ref'];
		delete workSheet['!rows'];

		defineGroups()



		// var group: string[][] = [];
		// for (let curKey in workSheet) {
		// 	let letter: number = 0, num: number = 0;
		// 	if (workSheet[curKey].w && workSheet[curKey].w.length < 10) {
		// 		for (let i = 0; i <= workSheet[curKey].w.length; i++) {
		// 			if (pattern.test(workSheet[curKey].w[i])) {
		// 				letter++;
		// 			}
		// 			else if (patternNum.test(workSheet[curKey].w[i]))
		// 				num++;
		// 		}
		// 	}
		// 	if (letter != 0 && num != 0 && letter < 5 && num < 5) {
		// 		group.push([workSheet[curKey].w, curKey]);
		// 	}
		// }

		// group.map((el: any) => {
		// 	console.log(el[0])

		// 	let startNumVal: string = '', startStrVal: string = '';
		// 	for (let i = 0; i <= el[1].length; i++) {
		// 		if (patternEng.test(el[1][i])) {
		// 			startStrVal += el[1][i]
		// 		}
		// 		else if (patternNum.test(el[1][i])) {
		// 			startNumVal += el[1][i]
		// 		}
		// 		if (el[1].length - i === 1)
		// 			break
		// 	}
		// 	for (let i = +startNumVal + 1; i <= +startNumVal + 10; i++) {
		// 		let str: string | undefined = workSheet[startStrVal + i] ? workSheet[startStrVal + i].w : undefined;
		// 		let yes: number = alphabet.indexOf(startStrVal) + 1;
		// 		if ((str && str.length < 6) || (!str && workSheet[alphabet[yes] + i] && workSheet[alphabet[yes] + i].w.length > 6)) {
		// 			let newStr = alphabet.indexOf(startStrVal);
		// 			startStrVal = alphabet[newStr + 1];
		// 			i = +startNumVal;
		// 			// if (isMerged(startStrVal, i, workSheet['!merges']))
		// 			// 	console.log('ЯЧЕЙКА ЕДИНА')
		// 		} else {
		// 			if (str) {
		// 				console.log(str)
		// 			}
		// 			// if (isMerged(startStrVal, i, workSheet['!merges']))
		// 			// 	console.log('ЯЧЕЙКА ЕДИНА')

		// 		}
		// 	}
		// console.log(+startNumVal)
		// })

		console.log(allGroups)
	}
	return (
		<div className="App">
			<input type="file" onChange={e => handleChange(e)} />
		</div>
	);
}
export default App;
