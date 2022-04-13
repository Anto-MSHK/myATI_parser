import React from 'react';

var XLSX: any;
if (typeof require !== 'undefined')
	XLSX = require('xlsx');

type lesson = {
	lessonName: string
	teacherInfo: string
	cabinet: string
	time: string
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

var alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO'];

function App() {
	const isMerged = (str: string, num: number, merges: { s: { r: {}, c: {} }, e: { r: {}, c: {} } }[]) => {
		let v: boolean = false;
		merges.map(el => {
			if ((el.s.c === alphabet.indexOf(str) && el.e.c === alphabet.indexOf(str)) && (el.s.r === num && el.e.r === num + 1)) {
				v = true
			}
		})
		return v
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

		const workSheet: list = workBook.Sheets[workBook.SheetNames[0]]

		delete workSheet['!margins'];
		// delete workSheet['!merges'];
		delete workSheet['!ref'];
		delete workSheet['!rows'];

		var pattern = /^[А-Я]+$/;
		var patternEng = /^[A-Z]+$/i;
		var patternNum = /^[0-9]+$/;



		var group: string[][] = [];
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
				group.push([workSheet[curKey].w, curKey]);
			}
		}

		group.map((el: any) => {
			console.log(el[0])

			let startNumVal: string = '', startStrVal: string = '';
			for (let i = 0; i <= el[1].length; i++) {
				if (patternEng.test(el[1][i])) {
					startStrVal += el[1][i]
				}
				else if (patternNum.test(el[1][i])) {
					startNumVal += el[1][i]
				}
				if (el[1].length - i === 1)
					break
			}
			for (let i = +startNumVal + 1; i <= +startNumVal + 10; i++) {
				let str: string | undefined = workSheet[startStrVal + i] ? workSheet[startStrVal + i].w : undefined;
				let yes: number = alphabet.indexOf(startStrVal) + 1;
				if ((str && str.length < 6) || (!str && workSheet[alphabet[yes] + i] && workSheet[alphabet[yes] + i].w.length > 6)) {
					let newStr = alphabet.indexOf(startStrVal);
					startStrVal = alphabet[newStr + 1];
					i = +startNumVal;
					if (isMerged(startStrVal, i, workSheet['!merges']))
						console.log('ЯЧЕЙКА ЕДИНА')
				} else {
					// if (str) {
					console.log(startStrVal + i)
					// }
					if (isMerged(startStrVal, i, workSheet['!merges']))
						console.log('ЯЧЕЙКА ЕДИНА')

				}
			}
			// console.log(+startNumVal)
		})

		console.log(workSheet)
	}
	return (
		<div className="App">
			<input type="file" onChange={e => handleChange(e)} />
		</div>
	);
}

export default App;
