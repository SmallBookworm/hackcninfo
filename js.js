/**
 * Created by peng on 2017/7/31.
 */
var head = ['<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">',
	' <head>',
	'  <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>',
	'  <meta name="ProgId" content="Excel.Sheet"/>',
	'  <meta name="Generator" content="WPS Office ET"/>',
	'  <!--[if gte mso 9]>',
	'   <xml>',
	'    <o:DocumentProperties>',
	'     <o:Created>2017-08-01T14:50:16</o:Created>',
	'     <o:LastAuthor>peng</o:LastAuthor>',
	'     <o:LastSaved>2017-08-01T15:11:54</o:LastSaved>',
	'    </o:DocumentProperties>',
	'    <o:CustomDocumentProperties>',
	'     <o:KSOProductBuildVer dt:dt="string">2052-10.1.0.6690</o:KSOProductBuildVer>',
	'    </o:CustomDocumentProperties>',
	'   </xml>',
	'  <![endif]-->',
	'  <!--[if gte mso 9]>',
	'   <xml>',
	'    <x:ExcelWorkbook>',
	'     <x:ExcelWorksheets>',
	'      <x:ExcelWorksheet>',
	'       <x:Name>ok (1)</x:Name>',
	'       <x:WorksheetOptions>',
	'        <x:DefaultRowHeight>270</x:DefaultRowHeight>',
	'        <x:Selected/>',
	'        <x:Panes>',
	'         <x:Pane>',
	'          <x:Number>3</x:Number>',
	'          <x:ActiveCol>3</x:ActiveCol>',
	'          <x:ActiveRow>6</x:ActiveRow>',
	'          <x:RangeSelection>D7</x:RangeSelection>',
	'         </x:Pane>',
	'        </x:Panes>',
	'        <x:DoNotDisplayGridlines/>',
	'        <x:ProtectContents>False</x:ProtectContents>',
	'        <x:ProtectObjects>False</x:ProtectObjects>',
	'        <x:ProtectScenarios>False</x:ProtectScenarios>',
	'        <x:Print>',
	'         <x:PaperSizeIndex>9</x:PaperSizeIndex>',
	'        </x:Print>',
	'       </x:WorksheetOptions>',
	'      </x:ExcelWorksheet>',
	'     </x:ExcelWorksheets>',
	'     <x:ProtectStructure>False</x:ProtectStructure>',
	'     <x:ProtectWindows>False</x:ProtectWindows>',
	'     <x:WindowHeight>13050</x:WindowHeight>',
	'     <x:WindowWidth>28695</x:WindowWidth>',
	'    </x:ExcelWorkbook>',
	'   </xml>',
	'  <![endif]-->',
	' </head>'
].join("");
var mfuck = [];
addS();

function addS() {
	var url = path;
	if (!hisFlag) {
		url += "/disclosure/" + column + "_" + tabName;
	} else {
		url += "/announcement/query";
	}
	var total = Math.ceil(66926 / 30);
	var throwAmount = 0;
	for (let pageNum = 1; pageNum <= total; pageNum++) {
		//pageNum += 1;
		$("#pageNum_hidden_input").val(pageNum);
		$.ajax({
			url: url,
			type: 'POST',
			data: $('#AnnoucementsQueryForm').serialize(),
			dataType: 'json',
			error: function () {
				dataLoadingProgress(tabName, 1);
			},
			success: function (result) {
				if (result != null) {
					for (var i of result.announcements) {
						if (i.announcementTitle.indexOf('摘要') != -1) {
							throwAmount++;
							continue;
						}
						var date = fomatDate(i.announcementTime);
						for (var l = 0; l < 6 - i.secCode.length; l++)
							i.secCode = '0' + i.secCode;
						var one = {
							代码: i.secCode,
							简称: i.secName,
							公告标题: i.announcementTitle,
							公告时间: date,
							年度报告网址: 'http://www.cninfo.com.cn/cninfo-new/disclosure/szse/download/' + i.announcementId + '?announceTime=' + date
						};
						mfuck.push(one);
					}

					if ((mfuck.length+throwAmount) == total * 30)
						downloadFile('ok.xls', head + '<body><table>' + createTable(mfuck).innerHTML + '</table></body></html>')
					else
						console.log('' + (((mfuck.length+throwAmount) / (total * 30) * 100).toFixed(2) + '%'));
				}
			}
		});
	}

	function createTable(json) {
		var ctable = document.createElement('table');
		ctable.innerHTML = '';
		init(json, ctable);

		for (var i in json) {
			var tr = document.createElement('tr');
			for (var j in json[i]) {
				var td = document.createElement('td');
				td.textContent = json[i][j];
				//			td.style.width = '100px';
				tr.appendChild(td);
			}
			ctable.appendChild(tr);
		}
		return ctable;
	}

	function init(json, ctable) {
		var thr = document.createElement('tr');
		for (var i in json[0]) {
			var th = document.createElement('th');
			th.textContent = i;
			thr.appendChild(th);
		}
		ctable.appendChild(thr);
	}

	function downloadFile(fileName, content) {
		var aLink = document.createElement('a');
		var blob = new Blob([content]);
		aLink.download = fileName;
		aLink.href = URL.createObjectURL(blob);
		aLink.click();
	}
}

var url = 'http://www.cninfo.com.cn/cninfo-new/disclosure/szse/download/1203746968?announceTime=2017-08-01';
var urlRange = [1, 1];
var hack = {};
hack.get = function (urlRange) {

}
hack.getUrl = function (urlRange) {

	return new Promise(function (resolve, reject) {
		var url = path;
		if (!hisFlag) {
			url += "/disclosure/" + column + "_" + tabName;
		} else {
			url += "/announcement/query";
		}
		var total = (urlRange[1] - urlRange[0] + 1) * 30;
		var finishAmount = 0;
		var failURL = [];
		for (let pageNum = urlRange[0]; pageNum <= urlRange[1]; pageNum++) {
			//pageNum += 1;
			$("#pageNum_hidden_input").val(pageNum);
			$.ajax({
				url: url,
				type: 'POST',
				data: $('#AnnoucementsQueryForm').serialize(),
				dataType: 'json',
				error: function () {
					dataLoadingProgress(tabName, 1);
				},
				success: function (result) {
					if (result != null) {
						for (var i of result.announcements) {
							var date = fomatDate(i.announcementTime);
							for (var l = 0; l < 6 - i.secCode.length; l++)
								i.secCode = '0' + i.secCode;
							var one = {
								代码: i.secCode,
								简称: i.secName,
								公告标题: i.announcementTitle,
								公告时间: date,
								年度报告网址:
							};
							if ()
								var url = 'http://www.cninfo.com.cn/cninfo-new/disclosure/szse/download/' + i.announcementId + '?announceTime=' + date;
							hack.downloadFile(url).then(function () {
								finishAmount++;
								console.log((finishAmount / (total) * 100).toFixed(2) + '%');
								if (finishAmount == total)
									console.log(failURL);
							}, function (info) {
								failURL.push(info);
								finishAmount++;
								console.log((finishAmount / (total) * 100).toFixed(2) + '%');
								if (finishAmount == total)
									console.log(failURL);
							})
						}

					}
				}
			});
		}

	});
}
hack.downloadFile = function (url) {
	var xhr = new XMLHttpRequest();

	if (xhr.response) {
		return new Promise(function (resolve, reject) {
			xhr.open('GET', url, true); //get请求，请求地址，是否异步
			xhr.responseType = "blob";

			xhr.onload = function () {
				if (xhr.status == 200) {
					var blob = xhr.response; // 注意:不是oReq.responseText
					if (blob) {
						var aLink = document.createElement('a');
						aLink.download = 'sss.PDF';
						aLink.href = URL.createObjectURL(blob);
						aLink.click();
						resolve(url);
					} else {
						reject(url);
					}
				} else {
					reject(url);
				}
			}
			xhr.onerror = function (e) {
				console.log(e);
				reject(url);
			}
			xhr.send();
		});

	} else {
		return new Promise(function (resolve, reject) {
			xhr.open('GET', url, false);
			//XHR binary charset opt by Marcus Granado 2006 [http://mgran.blogspot.com]
			xhr.overrideMimeType('text/plain; charset=x-user-defined');
			xhr.send(null);
			if (xhr.status != 200) {
				reject(url);
				console.log('error');
			}
			var n = xhr.responseText.length;
			var u8arr = new Uint8Array(n);
			while (n--) {
				u8arr[n] = xhr.responseText.charCodeAt(n);
			}
			var blob = new Blob([u8arr]);
			var aLink = document.createElement('a');
			aLink.download = 'sss.PDF';
			aLink.href = URL.createObjectURL(blob);
			aLink.click();
			resolve(url);
		});

	}

}

function loadImageAsync(url) {
	return new Promise(function (resolve, reject) {
		var image = new Image();

		image.onload = function () {
			resolve(image);
		};

		image.onerror = function () {
			reject(new Error('Could not load image at ' + url));
		};

		image.src = url;
	});
}