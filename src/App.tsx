import { type } from '@testing-library/user-event/dist/type';
import React from 'react';
import { createLessThan } from 'typescript';

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

type dayCells = {
	i_cell_row_last: number,
	i_cell_row_first: number
}

var alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ'];

var pattern = /^[А-Я]+$/;
var patternEng = /^[A-Z]+$/i;
var patternNum = /^[0-9]+$/;

var workSheet: list;

var allGroups: allGroups = {};

function App() {

	const checkingMerged = (str: string, num1: number, num2: number) => {
		let merges: merge[] = workSheet['!merges'];
		let isMergeCell: boolean = false;
		// num = -1
		merges.map(el => {
			if ((el.s.c === alphabet.indexOf(str) && el.e.c === alphabet.indexOf(str)) && (el.s.r === num1 - 1 && el.e.r === num2 - 1)) {
				isMergeCell = true
			}
		})
		return isMergeCell
	}
	function checkingGroupCellIsCorrect<N extends number, T extends string>(currentColumn: T, currentRow: N): T {
		let currentValue: string | undefined = workSheet[currentColumn + currentRow] ? workSheet[currentColumn + currentRow].w : undefined

		let nextCell: number = alphabet.indexOf(currentColumn) + 1;
		let incorrectData = (currentValue && currentValue.length < 6);
		let emptyData = (!currentValue && workSheet[alphabet[nextCell] + currentRow] && workSheet[alphabet[nextCell] + currentRow].w.length > 6);

		if (incorrectData || emptyData) {
			return checkingGroupCellIsCorrect(alphabet[nextCell] as T, currentRow)
		} else {
			return currentColumn
		}
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
				let correctColumn = checkingGroupCellIsCorrect(referСell_letter, +referСell_number + 1)
				referСell_letter = correctColumn ? correctColumn : referСell_letter
				allGroups[workSheet[curKey].w] = defineSchedule(referСell_letter, referСell_number, workSheet[curKey].w);
			}
		}
	}

	const defineSchedule = (referСell_letter: string, referСell_number: string, group: string) => {
		var stydyWeek: stydyWeek = {
			_debagProps: { _referenceKey: undefined, _referenceKey_isСorrect: undefined },
			days: [undefined, undefined, undefined, undefined, undefined, undefined]
		};
		let iCurrentWeekDay = 0;

		let i_letter = alphabet.indexOf(referСell_letter);

		let cellsOfDays: dayCells[] = [{ i_cell_row_first: 0, i_cell_row_last: 0 }];
		let i_day = 0;
		let i_cell_row = +referСell_number + 1;
		let i_cell_row_last = +referСell_number + 20;

		for (let i_cell_column = i_letter - 1; i_cell_column >= 0; i_cell_column--) {
			while (i_cell_row_last > i_cell_row) {
				for (let i_cell_row_first = i_cell_row; i_cell_row_first <= i_cell_row + 5; i_cell_row_first++) {
					var isMerged = checkingMerged(alphabet[i_cell_column], i_cell_row_first, i_cell_row_last),
						isCellLong = i_cell_row_last - i_cell_row_first > 6,
						isCellExists = workSheet[alphabet[i_cell_column] + i_cell_row_first] != undefined,
						isNotMD = isCellExists && workSheet[alphabet[i_cell_column] + i_cell_row_first].w != "ВОЕННАЯ КАФЕДРА"

					if (isMerged && isCellLong && isCellExists && isNotMD) {
						cellsOfDays[i_day] = { i_cell_row_first, i_cell_row_last }; i_day++
						i_cell_row = i_cell_row_last + 1
						i_cell_row_last += 20
						break
					}
				}
				i_cell_row_last--;
			}
			i_cell_row_last = +referСell_number + 20
		}

		cellsOfDays.map((day: dayCells) => {
			stydyWeek.days[iCurrentWeekDay] = defineDay(day, referСell_letter, referСell_number, group, iCurrentWeekDay)
			iCurrentWeekDay++;
		})

		return stydyWeek;
	}

	const defineDay = (rangeCells: dayCells, referСell_letter: string, referСell_number: string, group: string, weekDay: number) => {
		var stydyDay: stydyDay = {
			lessons: [undefined, undefined, undefined, undefined, undefined,]
		};

		let iCurrentLesson = 0;

		let column_lesson = alphabet[alphabet.indexOf(referСell_letter) - 1], column_cabinet = alphabet[alphabet.indexOf(referСell_letter) + 1]
		let cellsOfLesson: dayCells[] = [{ i_cell_row_first: 0, i_cell_row_last: 0 }], i_lesson = 0
		for (let i_cell = rangeCells.i_cell_row_first; i_cell <= rangeCells.i_cell_row_last; i_cell++) {
			for (let i_cell_last = i_cell + 1; i_cell_last <= i_cell + 5; i_cell_last++)
				if (checkingMerged(column_lesson, i_cell, i_cell_last)) {
					cellsOfLesson[i_lesson] = { i_cell_row_first: i_cell, i_cell_row_last: i_cell_last }
					i_lesson++
					break
				}
		}

		cellsOfLesson.map((lesson: dayCells) => {
			stydyDay.lessons[iCurrentLesson] = defineLesson(lesson, referСell_letter)
			iCurrentLesson++;
		})
		return stydyDay
	}

	const defineLesson = (referLesson: dayCells, referСell_letter: string) => {
		let lesson: lesson = { _debagProps: { _fullСontentСell: '' }, data: undefined };
		// referLesson += 1;


		for (let i_lesson = referLesson.i_cell_row_first; i_lesson <= referLesson.i_cell_row_last; i_lesson++) {
			let currentValue: string | undefined = workSheet[referСell_letter + i_lesson] ? workSheet[referСell_letter + i_lesson].w : undefined;
			if (checkingMerged(referСell_letter, i_lesson, referLesson.i_cell_row_last)) {
				if (currentValue)
					lesson._debagProps._fullСontentСell = currentValue + 'merge'
				else
					lesson._debagProps._fullСontentСell = 'void merge'
				break
			} else {
				if (currentValue) {
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
