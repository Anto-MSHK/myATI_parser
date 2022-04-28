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
type lessonInfo = {
	title?: string
	type?: string
	teacherInfo?: {
		name: string
		degree: string
	}
	cabinet?: string
}
type lesson = {
	_debagProps: {
		_fullСontentСell: string | undefined;
	}
	data: {
		topWeek: lessonInfo | undefined
		lowerWeek?: lessonInfo | undefined
		time?: string
	}
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

			if (workSheet[curKey].w && workSheet[curKey].w.length < 15) {
				for (let i = 0; i <= workSheet[curKey].w.length; i++) {
					if (pattern.test(workSheet[curKey].w[i])) {
						letter++;
					}
					else if (patternNum.test(workSheet[curKey].w[i]))
						num++;
				}
			}

			if (letter != 0 && num != 0 && letter <= 5 && num <= 5) {
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
				allGroups[workSheet[curKey].w] = defineSchedule(referСell_letter, referСell_number);
			}
		}
	}

	const defineSchedule = (referСell_letter: string, referСell_number: string) => {
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
			stydyWeek.days[iCurrentWeekDay] = defineDay(day, referСell_letter)
			iCurrentWeekDay++;
		})

		return stydyWeek;
	}

	const defineDay = (rangeCells: dayCells, referСell_letter: string) => {
		var stydyDay: stydyDay = {
			lessons: [undefined, undefined, undefined, undefined, undefined,]
		};

		let iCurrentLesson = 0;

		let column_lesson = alphabet[alphabet.indexOf(referСell_letter) - 1]
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

	const getOtherInfo = (str: string): [nameTeacher: string, degree: string, typeLesson: string] => {
		let surname = /^[А-Я][а-я]{1,20}\s[А-Я]\.[А-Я]\.$/
		let nameTeacher: string = ''
		let degree: string = ''
		let typeLesson: string = ''

		let i_degree: number = 0, i_type_lesson: number = 0;

		for (let i_char_start = 0; i_char_start <= str.length; i_char_start++) {
			let stop: boolean = false

			for (let i_char_end = str.length; i_char_end >= 1; i_char_end--) {
				let currentPhrase: string = '';

				for (let i = i_char_start; i <= i_char_end; i++) {
					currentPhrase += str[i]
				}
				if (surname.test(currentPhrase)) {
					nameTeacher = currentPhrase
					i_degree = i_char_start
					i_type_lesson = i_char_end
					stop = true
				}
			}

			if (stop === true)
				break
		}

		for (let i = 0; i <= i_degree - 1; i++) {
			degree += str[i]
		}

		for (let i = i_type_lesson + 1; i <= str.length; i++) {
			str[i] && (typeLesson += str[i])
		}

		return [nameTeacher, degree.trim(), typeLesson.trim()]
	}

	const getAllInfo = (str: string): [title: string, nameTeacher: string, degree: string, typeLesson: string] => {
		let surname = /^[А-Я][а-я]{1,20}\s[А-Я]\.[А-Я]\.$/

		let title: string = ''
		let nameTeacher: string = ''
		let degree: string = ''
		let typeLesson: string = ''

		let i_degree: number = 0, i_type_lesson: number = 0;

		for (let i_char_start = 0; i_char_start <= str.length; i_char_start++) {
			let stop: boolean = false

			for (let i_char_end = str.length; i_char_end >= 1; i_char_end--) {
				let currentPhrase: string = '';

				for (let i = i_char_start; i <= i_char_end; i++) {
					currentPhrase += str[i]
				}
				if (surname.test(currentPhrase)) {
					nameTeacher = currentPhrase
					i_degree = i_char_start
					i_type_lesson = i_char_end
					stop = true
				}
			}

			if (stop === true)
				break
		}

		let i_first_t = str.indexOf('.');
		if (i_first_t != -1) {
			for (let j = i_first_t; j >= 0; j--) {
				if (str[j] === ' ') {
					i_first_t = j;
					break
				}
			}
			for (let i = 0; i <= i_first_t; i++) {
				title += str[i]
			}
		}

		for (let i = i_first_t - 1; i <= i_degree - 1; i++) {
			degree += str[i]
		}

		for (let i = i_type_lesson + 1; i <= str.length; i++) {
			str[i] && (typeLesson += str[i])
		}

		return [title, nameTeacher, degree.trim(), typeLesson.trim()]
	}

	const defineLesson = (referLesson: dayCells, referСell_letter: string) => {
		let lesson: lesson = { _debagProps: { _fullСontentСell: '' }, data: { topWeek: {} } };
		// referLesson += 1;
		let column_cabinet = alphabet[alphabet.indexOf(referСell_letter) + 1]

		let currentLesson: string | undefined = workSheet[referСell_letter + referLesson.i_cell_row_first] ? workSheet[referСell_letter + referLesson.i_cell_row_first].w : undefined;
		let currentTeacher: string | undefined = workSheet[referСell_letter + referLesson.i_cell_row_last] ? workSheet[referСell_letter + referLesson.i_cell_row_last].w : undefined;

		if (checkingMerged(column_cabinet, referLesson.i_cell_row_first, referLesson.i_cell_row_last) && (currentLesson && currentTeacher)) {
			let currentCabinet: string | undefined = workSheet[column_cabinet + referLesson.i_cell_row_first] ? workSheet[column_cabinet + referLesson.i_cell_row_first].w : undefined;


			let otherData = currentTeacher ? getOtherInfo(currentTeacher) : [];
			lesson.data.topWeek && (lesson.data.topWeek = {
				title: currentLesson,
				cabinet: currentCabinet,
				teacherInfo: { name: otherData[0], degree: otherData[1] },
				type: otherData[2]
			})

		} else {
			let currentCabinetTop: string | undefined = workSheet[column_cabinet + referLesson.i_cell_row_first] ? workSheet[column_cabinet + referLesson.i_cell_row_first].w : undefined;
			let currentCabinetLower: string | undefined = workSheet[column_cabinet + referLesson.i_cell_row_last] ? workSheet[column_cabinet + referLesson.i_cell_row_last].w : undefined;

			let currentDataTop: string | undefined = workSheet[referСell_letter + referLesson.i_cell_row_first] ? workSheet[referСell_letter + referLesson.i_cell_row_first].w : undefined;
			let currentDataLower: string | undefined = workSheet[referСell_letter + referLesson.i_cell_row_last] ? workSheet[referСell_letter + referLesson.i_cell_row_last].w : undefined;

			let dataTop = currentDataTop ? getAllInfo(currentDataTop) : [];
			let dataLower = currentDataLower ? getAllInfo(currentDataLower) : [];

			if (currentDataTop)
				lesson.data.topWeek = {
					title: dataTop[0],
					type: dataTop[3],
					teacherInfo: {
						name: dataTop[1],
						degree: dataTop[2]
					}
				}
			else
				lesson.data.topWeek = undefined

			if (currentDataLower)
				lesson.data.lowerWeek = {
					title: dataLower[0],
					type: dataLower[3],
					teacherInfo: {
						name: dataLower[1],
						degree: dataLower[2]
					}
				}
			else
				lesson.data.lowerWeek = undefined

			lesson.data.topWeek && (lesson.data.topWeek.cabinet = currentCabinetTop ? currentCabinetTop : undefined);
			lesson.data.lowerWeek && (lesson.data.lowerWeek.cabinet = currentCabinetLower ? currentCabinetLower : undefined);
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

		console.log(allGroups)
	}
	return (
		<div className="App">
			<input type="file" onChange={e => handleChange(e)} />
		</div>
	);
}
export default App;
