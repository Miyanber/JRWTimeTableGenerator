﻿'use strict';

(function () {
	class loadSpreadSheet {
		constructor() {
			this.DESTINATION_URL = "https://script.google.com/macros/s/AKfycbwlrB8oXEt4hLlEjsQkVQ-tVjHdbVHRHzHeLryZnOq2zMzn_To6ymBqTW0QWfibbNk/exec";
			this.TRAIN_TYPE_URL = "https://script.google.com/macros/s/AKfycbz2YGPFbOk5RqOju0AONsqDU6AnRzA-X-hORI4qJBM7sjZErZHsvspHSIYPieW3SEkyqQ/exec";
		}
		async getArray(url) {
			return await fetch(url).then((response) => {
				return response.json();
			}).then((json) => {
				return JSON.parse(JSON.stringify(json, null, " "));
			}).then((json) => {
				return json.allData;
			}).catch((error) => {
				console.log(error);
				throw Error(`Cannot load the Spread Sheet. URL:'${url}'.`);
			});
		}
		async getDestinationArray() {
			return await this.getArray(this.DESTINATION_URL);
		}
		async getTrainTypeArray() {
			return await this.getArray(this.TRAIN_TYPE_URL);
		}
	}
	let copyState, worksheetW, worksheetH;
	Office.onReady(function () {
		// Office is ready
		const arrFunc = [];
		const loadSpreadSheetConstructor = new loadSpreadSheet();
		let destinationArray, trainTypeArray;
		arrFunc.push(async (resolve) => { destinationArray = await loadSpreadSheetConstructor.getDestinationArray(); resolve(); });
		arrFunc.push(async (resolve) => { trainTypeArray = await loadSpreadSheetConstructor.getTrainTypeArray(); resolve(); });
		const arrPromise = arrFunc.map((func) => new Promise(func));
		
		$(document).ready(function () {
			// The document is ready
			Promise.all(arrPromise).then(() => {
				$("#loading").addClass("loaded");
			});
			copyState = $("#copyState");
			worksheetW = $("#worksheetW");
			worksheetH = $("#worksheetH");
			let fileW = document.getElementById("file1");
			let fileWState = [$("#weekdaysFileStateEki"), $("#weekdaysFileStateDir"), $("#weekdaysFileStateDay"),];
			let fileH = document.getElementById("file2");
			let fileHState = [$("#holidaysFileStateEki"), $("#holidaysFileStateDir"), $("#holidaysFileStateDay"),];
			fileW.addEventListener("change", () => {
				eventListener(document.getElementById("file1"), fileWState, false);
			});
			fileH.addEventListener("change", () => {
				eventListener(document.getElementById("file2"), fileHState, true);
			})
			/**
			 * 
			 * @param {HTMLInputElement} file
			 * @param {Array<JQuery>} fileState
			 * @param {Boolean} isHoliday1
			 */
			function eventListener(file, fileState, isHoliday1) {
				fileState[2].css("color", "black");
				fileState[2].css("font-weight", "normal");
				fileState[0].text("選択後表示");
				fileState[1].text("選択後表示");
				fileState[2].text("選択後表示");
				let files = file.files[0];
				if (files.length == 0) {
					return
				}
				const reader = new FileReader();
				reader.onload = () => {
					const HTML = document.createElement("section");
					document.body.appendChild(HTML);
					HTML.style.display = "none";
					HTML.innerHTML = reader.result;
					if (document.getElementsByName("EKI")[0] == null) {
						fileState.text("不正なファイルです。");
						return
					}
					let dateString = String(document.getElementById("SearchDate").value);//20220730
					let date = ((dateString) => {
						const year = Number(dateString.substr(0, 4)),
							month = Number(dateString.substr(4, 2)),
							today = Number(dateString.substr(6, 2));
						let date = new Date(year, month - 1, today);//month is 0-index.
						return date;
					})(dateString);
					let isHoliday2 = isHolidays(date);
					if (isHoliday1 == isHoliday2) {
						if (isHoliday2) {
							fileState[2].text("土・休日用");
						} else {
							fileState[2].text("平日用");
						}
					} else {
						fileState[2].css("color", "red");
						fileState[2].css("font-weight", "bold");
						if (isHoliday2) {
							fileState[2].text("土・休日用");
						} else {
							fileState[2].text("平日用");
						}
					}
					fileState[0].text(document.getElementsByName("EKI")[0].value + "駅");
					fileState[1].text(document.getElementsByName("DIR")[0].value);
					document.body.removeChild(HTML);
				};
				reader.readAsText(files);
			}

			/**
			 * 
			 * @param {Date} date
			 */
			function isHolidays(date) { //同名の変数はローカル変数が優先される
				if (date.getDay() == 0 || date.getDay() == 6) {
					return true; //土・日
				} else {
					for (let i = 0; i < holidays.length; i++) {
						if (date.getFullYear() == holidays[i][0] && date.getMonth() + 1 == holidays[i][1] && date.getDate() == holidays[i][2]) {
							return true;
						}
					}
				}
				return false
			}

			async function run() {
				message("start", "作成中...");
				const trainAllData = await getAllData();
				if (trainAllData == null) {
					return
				}
				console.log(trainAllData);
				const usage = trainAllData[2],
					station = trainAllData[0][1],
					weekdays = trainAllData[0],
					holidays = trainAllData[1];
				//baseシートの存在確認
				try {
					await Excel.run(async (context) => {
						let basedSheet = context.workbook.worksheets.getItem("base");
						await context.sync();
					})
                } catch (e) {
					message("error", `「base」シートが存在しません。<br>ダウンロードしてください。`);
					console.log(e);
					return;
                }
				await Excel.run(async function (context) {
					// weekdays,holidaysの中身：{ trainData, station, sheetOptions }
					let basedSheet = context.workbook.worksheets.getItem("base");
					let sheet = basedSheet.copy(Excel.WorksheetPositionType.beginning);
					const worksheetUsage = `${station.eki}-${station.dir}`;
					sheet.name = worksheetUsage;
					sheet.getRange("C44").values = [
						[usage]
					];
					sheet.getRange("G2").values = [
						[station.eki]
					]; //駅
					sheet.getRange("P2").values = [
						[station.dir]
					]; //方面
					sheet.getRange("AQ2").values = [
						[station.eki]
					]; //駅
					sheet.getRange("AZ2").values = [
						[station.dir]
					]; //方面
					//#region ここからシート調整
					let columnRange = sheet.getRange("B2:B44");
					let count;
					columnRange.getOffsetRange(0, 1).getColumnsAfter(weekdays[2].deletedColumns[0]).delete("Left");
					count = weekdays[2].deletedColumns[0];
					columnRange.getOffsetRange(0, 5 - count).getColumnsAfter(weekdays[2].deletedColumns[1]).delete("Left");
					count += weekdays[2].deletedColumns[1];
					columnRange.getOffsetRange(0, 14 - count).getColumnsAfter(weekdays[2].deletedColumns[2]).delete("Left");
					count += weekdays[2].deletedColumns[2];
					columnRange.getOffsetRange(0, 26 - count).getColumnsAfter(weekdays[2].deletedColumns[3]).delete("Left");
					count += weekdays[2].deletedColumns[3];
					columnRange.getOffsetRange(0, 37 - count).getColumnsAfter(holidays[2].deletedColumns[0]).delete("Left");
					count += holidays[2].deletedColumns[0];
					columnRange.getOffsetRange(0, 41 - count).getColumnsAfter(holidays[2].deletedColumns[1]).delete("Left");
					count += holidays[2].deletedColumns[1];
					columnRange.getOffsetRange(0, 50 - count).getColumnsAfter(holidays[2].deletedColumns[2]).delete("Left");
					count += holidays[2].deletedColumns[2];
					columnRange.getOffsetRange(0, 62 - count).getColumnsAfter(holidays[2].deletedColumns[3]).delete("Left");
					//#endregion
					let rangeW = sheet.getRange("C3"),
						rangeH = rangeW.getCell(0, weekdays[2].number + 3);
					setAllData(rangeW, weekdays, true, station);
					setAllData(rangeH, holidays, false, station);
					await context.sync();
					message("ok", "処理完了");
				});
				/*await Excel.run(async (context) => {
					context.workbook.worksheets.getFirst().name = worksheetUsage;
					await context.sync();
                })*/
			}

			/**
			 * 
			 * @param {Excel.Range} range
			 * @param {Array} data
			 * @param {Boolean} isWeekday
			 * @param {Object} station
			 */
			function setAllData(range, data, isWeekday, station) {
				const sheetNumber = data[2].number,
					columnWidth = data[2].columnWidth,
					columnWidthHour = data[2].columnWidthHour,
					widthBetween = data[2].widthBetween;
				const trainAllData = data[0],
					obj = data[1];
				range.getOffsetRange(0, -1).getColumnsAfter(sheetNumber).set({
					format: {
						columnWidth: columnWidth
					}
				});
				range.getOffsetRange(0, -1).getEntireColumn().getColumn(0).set({
					format: {
						columnWidth: columnWidthHour
					}
				});
				range.getOffsetRange(0, sheetNumber - 1).getColumnsAfter(2).set({
					format: {
						columnWidth: widthBetween
					}
				});
				for (var i = 4; i <= 25; i++) {
					if (trainAllData[i] == null) {
						continue;
					}
					let limitedExpress = 0;
					for (var l = 0; l < trainAllData[i].length; l++) {
						let obj = trainAllData[i][l];
						if (obj.isLimitedExpress && station.includesLimitedExp) {
							limitedExpress += 1;
							continue;
						}
						let row = 2 * Number(obj.hour) - 9, //行 YOKOに向かっているから縦のやつ！！
							column = obj.index - 1 - limitedExpress; //列 TATEに向かっているから横のやつ！！
						if (obj.hour == 4) {
							const hourCell = range.getCell(0, sheetNumber - 5);
							hourCell.values = [
								[4]
							];
							if (isWeekday) {
								hourCell.format.fill.color = "#0640a4";
							} else {
								hourCell.format.fill.color = "red";
							}
							hourCell.format.font.color = "white";
							row = 2 * Number(5) - 9;
							column = obj.index - 1 + sheetNumber - 4;
						} else if (obj.hour == 25) {
							const hourCell = range.getCell(38, sheetNumber - 5);
							hourCell.values = [
								[1]
							];
							if (isWeekday) {
								hourCell.format.fill.color = "#0640a4";
							} else {
								hourCell.format.fill.color = "red";
							}
							hourCell.format.font.color = "white";
							row = 2 * Number(24) - 9;
							column = obj.index - 1 + sheetNumber - 4;
						}
						let cellT = range.getCell(row, column),
							cellF = range.getCell(row - 1, column);
						cellT.values = [
							[obj.minutes]
						];
						if (obj.minutes >= 10 && sheetNumber >= 18 && sheetNumber <= 21) {
							cellT.format.font.name = "Arial Narrow";
						} else {
							cellT.format.font.name = "Arial";
						}
						cellT.values = [
							[obj.minutes]
						];
						cellF.values = [
							[obj.trainTypeSmall + obj.destinationSmall]
						];
						if (obj.trainTypeColor != "") {
							cellT.format.font.color = obj.trainTypeColor;
						}
						if (obj.trainTypeFill != "") {
							cellT.format.fill.color = obj.trainTypeFill;
						}
						if (obj.destinationColor != "") {
							if (obj.destinationColor.split(",")[1] != undefined) {
								if (station.senk.includes(obj.destinationColor.split(",")[1])) {
									cellF.format.font.color = obj.destinationColor.split(",")[0];
								} else { }
							} else {
								cellF.format.font.color = obj.destinationColor;
							}
						}
					}
				}
			}

			async function getAllData() {
				const json = [];
				let destinationUsageArray = [], trainTypeUsageArray = [];
				fileW = document.getElementById("file1").files[0];
				fileH = document.getElementById("file2").files[0];
				if (fileW == null || fileH == null) {
					message("error", "ファイルを選択してください!");
					return null;
				}
				const readerW = new FileReader();
				const readerH = new FileReader();
				//平日
				await (async() => {
					readerW.readAsText(fileW);
					await new Promise(resolve => readerW.onload = () => resolve());
					let element = getTrainAllData(readerW.result, "weekdays");
					destinationUsageArray = [...destinationUsageArray, ...element[1].usage[0]];
					trainTypeUsageArray = [...trainTypeUsageArray, ...element[1].usage[1]];
					json.push(element);
				})();
				//土・休日
				await (async() => {
					readerH.readAsText(fileH);
					await new Promise(resolve => readerH.onload = () => resolve());
					let element = getTrainAllData(readerH.result, "holidays");
					destinationUsageArray = [...destinationUsageArray, ...element[1].usage[0]];
					trainTypeUsageArray = [...trainTypeUsageArray, ...element[1].usage[1]];
					json.push(element);
				})();
				//#region Usage作成
				destinationUsageArray = [...new Set(destinationUsageArray)].sort((a, b) => {
					if (b.includes("無印＝")) {
						return -1;
					} else if (a.includes("無印＝")) {
						return 1;
					} else {
						return 0;
					}
				});
				trainTypeUsageArray = trainTypeUsageArray.filter((value, index, array) => array.findIndex((dataElement) => dataElement[0] === value[0]) === index //これで二つ目以降の要素を排除できる！
				).sort((a, b) => {
					return b[1] - a[1];
				});
				const trainTypeUsageArray1 = trainTypeUsageArray.filter((value) => value[1] >= 20).map((value) => value[0]);
				const trainTypeUsageArray2 = trainTypeUsageArray.filter((value) => value[1] < 20).map((value) => value[0]);
				let usage = `${trainTypeUsageArray1.join("　")}\n${trainTypeUsageArray2.join("　")}\n${destinationUsageArray.join("　")}`;
				if (usage.includes("\n\n")) {
					usage.replace("\n\n", "\n");
				}
				//#endregion
				json.push(usage);
				return json
			}

			/**
			 * 
			 * @param {String} innerHTML HTML String
			 * @param {String} dayString "weekdays" or "holidays"
			 * @returns {Array} [trainData,station,sheetOption]
			 */
			function getTrainAllData(innerHTML, dayString) {
				let json = [];
				//Web時刻表の要素を作成
				const HTML = document.createElement("section");
				document.body.appendChild(HTML);
				HTML.style.display = "none";
				HTML.innerHTML = innerHTML;
				let station = {
					eki: "",
					dir: "",
					day: "",
					senk: "",
					includesLimitedExp: false,
					usage: null
				},
					trainData = [];
				if (document.getElementsByName("EKI")[0] == null) {
					message("error", "データの取得に失敗しました。<br>URLが不正です。");
					throw Error("Can't get data from the URL.");
				}
				station.includesLimitedExp = document.getElementById("limited").checked;
				station.eki = document.getElementsByName("EKI")[0].value;
				station.dir = document.getElementsByName("DIR")[0].value;
				station.day = dayString; //平日/土・休日
				station.senk = document.getElementsByName("SENK")[0].value;
				document.querySelectorAll(".hour").forEach((element) => {
					const tr = element.parentNode, //<tr>
						nbsp = tr.lastElementChild; //&nbsp判定用
					let hour = parseInt(tr.firstChild.innerHTML); //<th>
					if (hour <= 2) {
						hour += 24; //時刻調整
					}
					if (element.rowSpan == 1) {
						//時間(hour)ごとに区切ってる
						trainData[hour] = search(nbsp, tr, false, hour, station.eki, station.senk);
					} else if (element.rowSpan == 2) {
						const tr2 = tr.nextElementSibling,
							nbsp2 = tr2.lastElementChild;
						trainData[hour] = [...search(nbsp, tr, false, hour, station.eki, station.senk), ...search(nbsp2, tr2, true, hour, station.eki, station.senk)]; //配列の結合
					} else {
						console.log("error:element.rowSpan > 2");
					}
				});
				station.eki += "駅";
				let date = new Date();
				//ここから各種シート設定
				let maxIndex = [],
					destinationUsageArray = [],
					trainTypeUsageArray = [];
				//シートの最大値の作成・凡例の配列の作成
				trainData.forEach((value) => { //valueは時間ごとのtrainData
					if (value === null) {
						return
					}
					let valueLength = 0;
					valueLength = value.length;
					value.forEach((trainData) => {
						if (trainData === null) {
							return
						}
						if (trainData.isLimitedExpress && station.includesLimitedExp) {
							valueLength += -1;
						}
						if (!station.includesLimitedExp) { //特急列車を除外しない場合
							destinationUsageArray.push(trainData.destinationUsage);
							trainTypeUsageArray.push(trainData.trainTypeUsage);
						} else {
							if (!trainData.isLimitedExpress) { //特急列車を除外し、該当列車が特急列車ではない場合
								destinationUsageArray.push(trainData.destinationUsage);
								trainTypeUsageArray.push(trainData.trainTypeUsage);
							}
						}
					});
					maxIndex.push(valueLength);
				});
				station.usage = [destinationUsageArray, trainTypeUsageArray];
				const MAX_INDEX = Math.max(...maxIndex);
				console.log(MAX_INDEX);
				let sheetOptions = {
					number: 33,
					columnWidth: conversion(11 / 33),
					columnWidthHour: conversion(0.5),
					widthBetween: conversion(0.3),
					deletedColumns: [0, 0, 0, 0] //元は[4,9,12,8]
				}
				if (MAX_INDEX > 30) {
					sheetOptions.number = 33;
					sheetOptions.columnWidth = conversion(11 / sheetOptions.number);
					sheetOptions.deletedColumns = [0, 0, 0, 0];
				} else if (MAX_INDEX > 28) {
					sheetOptions.number = 30;
					sheetOptions.columnWidth = conversion(11 / sheetOptions.number);
					sheetOptions.deletedColumns = [1, 0, 1, 1];
				} else if (MAX_INDEX > 25) {
					sheetOptions.number = 28;
					sheetOptions.columnWidth = conversion(11 / sheetOptions.number);
					sheetOptions.deletedColumns = [1, 1, 1, 2];
				} else if (MAX_INDEX > 23) {
					sheetOptions.number = 25;
					sheetOptions.columnWidth = conversion(11 / sheetOptions.number);
					sheetOptions.deletedColumns = [1, 3, 2, 2];
				} else if (MAX_INDEX > 21) {
					sheetOptions.number = 23;
					sheetOptions.columnWidth = conversion(11 / sheetOptions.number);
					sheetOptions.deletedColumns = [1, 4, 2, 3];
				} else if (MAX_INDEX > 18) {
					sheetOptions.number = 21;
					sheetOptions.columnWidth = conversion(5.15 / sheetOptions.number);
					sheetOptions.columnWidthHour = conversion(0.4),
						sheetOptions.widthBetween = conversion(0.2);
					sheetOptions.deletedColumns = [1, 5, 3, 3];
				} else if (MAX_INDEX > 15) {
					sheetOptions.number = 18;
					sheetOptions.columnWidth = conversion(5.15 / sheetOptions.number);
					sheetOptions.columnWidthHour = conversion(0.4),
						sheetOptions.widthBetween = conversion(0.2);
					sheetOptions.deletedColumns = [1, 5, 6, 3];
				} else if (MAX_INDEX > 12) {
					sheetOptions.number = 15;
					sheetOptions.columnWidth = conversion(5.15 / sheetOptions.number);
					sheetOptions.columnWidthHour = conversion(0.4),
						sheetOptions.widthBetween = conversion(0.2);
					sheetOptions.deletedColumns = [1, 6, 7, 4];
				} else {
					sheetOptions.number = 12;
					sheetOptions.columnWidth = conversion(5.05 / sheetOptions.number);
					sheetOptions.columnWidthHour = conversion(0.4),
						sheetOptions.widthBetween = conversion(0.3);
					sheetOptions.deletedColumns = [2, 6, 9, 4];
				}
				json = [trainData, station, sheetOptions];
				worksheet(dayString).innerHTML = `<strong>${sheetOptions.number}</strong>以上`
				document.body.removeChild(HTML);
				return json;
				
			};
			/**
			 * 
			 * @param {HTMLElement} nbsp
			 * @param {ParentNode} collection 
			 * @param {Boolean} isNext 
			 * @param {Number} hour 
			 * @param {String} eki 
			 * @param {String} senk 
			 * @returns {Object} trainData
			 */
			function search(nbsp, collection, isNext, hour, eki, senk) {
				let condition;
				if (String(nbsp.innerHTML) == "&nbsp;" && !isNext) {
					condition = collection.children.length - 2;
				} else {
					condition = collection.children.length - 1;
				}
				let element, index;
				let trainData = [];
				for (let i = 0; i < condition; i++) {
					if (isNext) {
						element = collection.children[i]; //td
						index = i + 17;
					} else {
						element = collection.children[i + 1];
						index = i + 1;
					}
					let information = {
						trainType: null,
						trainTypeSmall: null,
						trainTypeColor: null,
						trainTypeFill: null,
						trainTypeUsage: null,
						destination: null,
						destinationSmall: null,
						destinationColor: null,
						destinationUsage: null,
						isLimitedExpress: false,
						hour: hour,
						minutes: String(element.lastElementChild /*<a> */.lastElementChild /*<font> */.lastElementChild /*<font> */.innerHTML),
						index: index
					};
					//行先判定
					for (let j = 0; j < destinationArray.length; j++) {
						//大阪環状線のみ初期設定
						if (senk.includes("大阪環状線")) {
							information.destination = "環状";
							information.destinationSmall = "";
							information.destinationColor = "";
							information.destinationUsage = `無印＝環状`;
						}
						const searchedText = String(element.firstChild.innerHTML), //Web時刻表表記
							searchText = destinationArray[j][0];
						if (searchText == "" || searchText == eki) {
							continue; //continueは、この回の処理をスキップするという意味。
						}
						if (searchedText.includes(searchText)) {
							const array = destinationArray[j];
							information.destination = array[0]; //Web時刻表表記
							information.destinationColor = array[2]; //小型時刻表表記の色
							let destinationSmallText = array[1].split(",");
							if (destinationSmallText[1] != undefined) {
								if (senk == destinationSmallText[1]) {
									information.destinationSmall = ""; //小型時刻表表記
								} else {
									information.destinationSmall = destinationSmallText[0];
								}
							} else {
								information.destinationSmall = destinationSmallText[0];
							}
							if (array[5] != false) {
								if (array[6] != false) { //「行」をつけるかどうか
									information.destinationUsage = `${information.destinationSmall}＝${array[5]}行`; //正式名称
									if (information.destinationSmall == "") {
										information.destinationUsage = `無印＝${array[5]}行`; //正式名称
									}
								} else {
									information.destinationUsage = `${information.destinationSmall}＝${array[5]}`; //正式名称
									if (information.destinationSmall == "") {
										information.destinationUsage = `無印＝${array[5]}`; //正式名称
									}
								}
							}
							//行先二つもあるのでもう一回
							if (array[4] != false) {
								for (j += 1; j < destinationArray.length; j++) {
									const searchedText = String(element.firstChild.innerHTML), //Web時刻表表記
										searchText = array[0];
									if (searchText == "" || searchText == eki) {
										continue; //continueは、この回の処理をスキップするという意味。
									}
									if (searchedText.includes(searchText)) {
										information.destination += " " + array[0]; //Web時刻表表記
										information.destinationSmall += array[1]; //小型時刻表表記
										information.destinationColor = array[2]; //小型時刻表表記の色
										break; //breakはforループごとスキップする。
									}
								}
							} else { }
							break; //breakはforループごとスキップする。
						}
					}
					//種別判定
					for (let j = 0; j < trainTypeArray.length; j++) {
						const searchedText = String(element.firstChild.innerHTML), //Web時刻表表記
							searchText = `>${trainTypeArray[j][0]}<`;
						if (searchText == "") {
							continue; //continueはこの回の処理をスキップする。
						}
						if (searchedText.includes(searchText)) {
							const array = trainTypeArray[j];
							if (array[5]) {
								information.isLimitedExpress = true;
							}
							information.trainType = array[0];
							information.trainTypeSmall = array[1]; //小型時刻表表記(▲など)
							information.trainTypeColor = array[2]; //小型時刻表表記の色
							information.trainTypeUsage = [array[3], array[6]]; //小型時刻表表記の説明,順番用index
							information.trainTypeFill = array[4]; //塗りつぶしの色
							break; //breakはforループごとスキップする。
						}
					}
					if (information.trainType == null) {
						information.trainType = "普通";
						information.trainTypeSmall = ""; //小型時刻表表記
						information.trainTypeColor = ""; //小型時刻表表記の色
						information.trainTypeFill = ""; //塗りつぶしの色
						information.trainTypeUsage = ["青字＝各駅に停車", 20];
					} else { }
					//ログ出力
					if (information.destination === null) {
						console.log(`行先が見つかりませんでした ${information.hour}時${information.minutes}分 種別:${information.trainType} index:${information.index}`);
						information.destination = "N"; //Web時刻表表記
						information.destinationSmall = "N"; //小型時刻表表記
						trainData.push(information);
						continue;
					} else if (information.minutes.includes("<i>")) {
						information.minutes = String(element.lastElementChild /*<a> */.lastElementChild /*<font> */.lastElementChild /*<font> */.lastElementChild /*<i> */.innerHTML);
						console.log(`時刻に変更有！${information.hour}時${information.minutes}分 種別:${information.trainType} 行先:${information.destination} index:${information.index}`);
						trainData.push(information);
						continue;
					} else {
						trainData.push(information);
						continue;
						console.log(`${information.hour}時${information.minutes}分 種別:${information.trainType} 行先:${information.destination} index:${information.index}`);
					}
				}
				return trainData;
			}

			function getTimeString() {
				let date = new Date();
				return `${date.getHours()}時${date.getMinutes()}分${date.getSeconds()}秒`
			}

			/**
			 * 
			 * @param {JQueryStatic} element
			 * @param {String} type
			 * @param {String} text
			 */
			function message(type, text = null) { //デフォルトは背景黄色
				const element = copyState;
				switch (type) {
					case "error":
						element.css("color", "white");
						element.css("background-color", "red");
						break;
					case "start":
						element.css("color", "black");
						element.css("background-color", "#e7e000");
						break;
					case "ok":
						element.css("color", "white");
						element.css("background-color", "#006d21");
						break;
					default:
						break;
				}
				if (text !== null) {
					element.html(text);
				}
			}
			/**
			 * 単位換算用関数
			 * @param {Number} number 変換する数値。
			 * @param {String} baseString もとの単位。デフォルトはcm。(px,ColumnWidth,mm,cm)
			 * @param {String} valueString 変換先の単位。デフォルトはColumnWidth。(px,ColumnWidth,mm,cm)
			 * @returns 
			 * 
			 */
			function conversion(number, baseString = "cm", valueString = "ColumnWidth") {
				let px;
				switch (baseString) { //pxに変換(小数点許容)
					case "px":
						px = number;
						break;
					case "ColumnWidth":
						px = number * 1.7;
						break;
					case "mm":
						px = number * 47.17 / 10;
						break;
					case "cm":
						px = number * 47.17;
						break;
				}
				px *= 4;
				switch (valueString) { //pxから変換
					case "px":
						return Math.round(px);
					case "ColumnWidth":
						return px * 10 / 17;
					case "mm":
						return px * 10 * 0.0212;
					case "cm":
						return px * 0.0212;
					default:
						return NaN
				}
			}

			function worksheet(day) {
				switch (day) {
					case "weekdays":
						return worksheetW;
					case "holidays":
						return worksheetH;
					case "reset":
						worksheetW.text("更新中...");
						worksheetH.text("更新中...");
						break;
					case "error":
						worksheetW.text("データ取得後に出力・更新されます");
						worksheetH.text("データ取得後に出力・更新されます");
						break;
				}
			}
			$("#run").click(async() => await run());
		});
	});
	/** Default helper for invoking an action and handling errors. */
	async function tryCatch(callback) {
		try {
			await callback();
		} catch (error) {
			// Note: In a production add-in, you'd want to notify the user through your add-in's UI.
			console.error(error);
		}
	}
	const holidays = [
		[2022, 1, 1],
		[2022, 1, 2],
		[2022, 1, 3],
		[2022, 1, 10],
		[2022, 2, 11],
		[2022, 2, 23],
		[2022, 3, 21],
		[2022, 4, 29],
		[2022, 4, 30],
		[2022, 5, 1],
		[2022, 5, 2],
		[2022, 5, 3],
		[2022, 5, 4],
		[2022, 5, 5],
		[2022, 7, 18],
		[2022, 8, 11],
		[2022, 8, 13],
		[2022, 8, 14],
		[2022, 8, 15],
		[2022, 8, 16],
		[2022, 9, 19],
		[2022, 9, 23],
		[2022, 10, 10],
		[2022, 11, 3],
		[2022, 11, 23],
		[2022, 12, 30],
		[2022, 12, 31],
		[2023, 1, 1],
		[2023, 1, 2],
		[2023, 1, 3],
		[2023, 1, 10],
		[2023, 2, 11],
		[2023, 2, 23],
		[2023, 3, 21],
		[2023, 4, 29],
		[2023, 4, 30],
		[2023, 5, 1],
		[2023, 5, 2],
		[2023, 5, 3],
		[2023, 5, 4],
		[2023, 5, 5],
		[2023, 7, 18],
		[2023, 8, 11],
		[2023, 8, 13],
		[2023, 8, 14],
		[2023, 8, 15],
		[2023, 8, 16],
		[2023, 9, 19],
		[2023, 9, 23],
		[2023, 10, 10],
		[2023, 11, 3],
		[2023, 11, 23],
		[2023, 12, 30],
		[2023, 12, 31]
	]
})();