const readEvents = (path, year) => {
	if (typeof require !== "undefined") XLSX = require("xlsx");
	const workbook = XLSX.readFile(path);

	const worksheetName = workbook.SheetNames[0];
	const worksheet = workbook.Sheets[worksheetName];

	const fetchCell = (worksheet, col, row) => worksheet[col + row];
	const letters = [
		"A",
		"B",
		"C",
		"D",
		"E",
		"F",
		"G",
		"H",
		"I",
		"J",
		"K",
		"L",
		"M",
		"N",
		"O",
		"P",
		"Q",
		"R",
		"S",
		"T",
		"U",
		"V",
		"W",
		"X",
		"Y",
		"Z",
	];

	const getNextLetter = (letter) => {
		let ind;
		if (letter.length == 2) {
			ind = [
				letters.indexOf(letter[0]) !== -1
					? letters.indexOf(letter[0])
					: undefined,
				letters.indexOf(letter[1]) !== -1
					? letters.indexOf(letter[1])
					: undefined,
			];
			if (ind[1] < letters.length - 1) {
				return letters[ind[0]] + letters[ind[1] + 1];
			} else if (ind[1] === letters.length - 1 && ind[0] < letters.length - 1) {
				return letters[ind[0] + 1] + letters[0];
			} else {
				return undefined;
			}
		} else if (letter.length == 1) {
			ind = letters.indexOf(letter) !== -1 ? letters.indexOf(letter) : undefined;
			if (ind < letters.length - 1) {
				return letters[ind + 1];
			} else if (ind === letters.length - 1) {
				return letters[0] + letters[0];
			} else {
				return undefined;
			}
		} else {
			return undefined;
		}
	};

	const fetchColumn = (worksheet, col) => {
		let index = 1;
		let currentCell = fetchCell(worksheet, col, index);

		const all = [];
		while (currentCell !== undefined) {
			all.push(currentCell);
			index += 1;
			currentCell = fetchCell(worksheet, col, index);
		}
		if (all.length === 0) {
			return undefined;
		}
		return all;
	};

	const fetchColumns = (worksheet) => {
		let index = "A";
		let currentColumn = fetchColumn(worksheet, index);

		const all = [];
		while (currentColumn !== undefined) {
			all.push(currentColumn);
			index = getNextLetter(index);
			currentColumn = fetchColumn(worksheet, index);
		}
		if (all.length === 0) {
			return undefined;
		}
		return all;
	};

	const indexedWorksheet = fetchColumns(worksheet);

	const getColumn = (worksheet, title) => {
		for (col of worksheet) {
			if (col[0].v === title) {
				return col.slice(1).map((row) => row.v);
			}
		}
	};

	const importantColumns = {
		date: getColumn(indexedWorksheet, "Date"),
		startHour: getColumn(indexedWorksheet, "Start"),
		endHour: getColumn(indexedWorksheet, "End"),
		note: getColumn(indexedWorksheet, "Notes"),
	};

	const getEvents = (cols) => {
		const categories = Object.keys(cols);
		const length = importantColumns[Object.keys(cols)[0]].length;
		const events = [];
		let event;
		for (let i = 0; i < length; i++) {
			event = {};
			for (category of categories) {
				event[category] = importantColumns[category][i];
			}
			events.push(event);
		}
		return events;
	};

	const events = getEvents(importantColumns);

	const parseHour = (hour) => {
		return [
			hour.substring(0, hour.indexOf(":")),
			hour.substring(hour.indexOf(":") + 1),
		];
	};

	const isBaNewDay = (a, b) => {
		const ap = parseHour(a);
		const bp = parseHour(b);
		const app = +ap[1] + +ap[0] * 60;
		const bpp = +bp[1] + +bp[0] * 60;
		if (app > bpp) {
			return true;
		}
		return false;
	};

	const parseEvent = (event) => {
		let newEvent = { note: event.note };
		const startDate = {
			day: event.date.substring(0, event.date.indexOf("/")),
			month: event.date.substring(event.date.indexOf("/") + 1),
		};
		const rDate = new Date(+year, +startDate.month - 1, +startDate.day);

		const rTDate = new Date(rDate.getTime());
		rTDate.setDate(rDate.getDate() + 1);

		let endDate;
		if (isBaNewDay(event.startHour, event.endHour)) {
			endDate = {
				day: rTDate.getDate(),
				month: rTDate.getMonth() + 1,
			};
		} else {
			endDate = {
				day: rDate.getDate(),
				month: rDate.getMonth() + 1,
			};
		}

		const startHour = parseHour(event.startHour);
		const endHour = parseHour(event.endHour);
		newEvent.startDate = new Date(
			+year,
			+startDate.month - 1,
			+startDate.day,
			+startHour[0],
			+startHour[1]
		);
		newEvent.endDate = new Date(
			+year,
			+endDate.month - 1,
			+endDate.day,
			+endHour[0],
			+endHour[1]
		);

		// newEvent.startDate = newEvent.startDate.toLocaleString();
		// newEvent.endDate = newEvent.endDate.toLocaleString();

		return newEvent;
	};

	const parsedEvents = events.map(parseEvent);

	return parsedEvents;
};

exports.readEvents = readEvents;