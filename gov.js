var downloadMoudle = {};
downloadMoudle.head = ['<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">',
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
downloadMoudle.createTable = function (json) {
	var ctable = document.createElement('table');
	ctable.innerHTML = '';
	downloadMoudle.init(json, ctable);

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
};
downloadMoudle.init = function (json, ctable) {
	var thr = document.createElement('tr');
	for (var i in json[0]) {
		var th = document.createElement('th');
		th.textContent = i;
		thr.appendChild(th);
	}
	ctable.appendChild(thr);
};
downloadMoudle.downloadFile = function (fileName, content) {
	var aLink = document.createElement('a');
	var blob = new Blob([content]);
	aLink.download = fileName;
	aLink.href = URL.createObjectURL(blob);
	aLink.click();
};

function Mata(name = "", type = "", info = {}, url = "", children = []) {
	this.name = name;
	this.type = type;
	this.url = url;
	this.children = children;
	this.info = info;
}
Mata.download = function (matas) {
	let arr = [];
	this.getArr(arr, matas);
	downloadMoudle.downloadFile('ok.xls', downloadMoudle.head + '<body><table>' + downloadMoudle.createTable(arr).innerHTML + '</table></body></html>');
}

Mata.getArr = function (arr, matas) {
	for (let i of matas) {
		let one = {
			type: i.type,
			name: i.name,
			"统计用区划代码": "",
			"城乡分类代码": ""

		};
		for (let j in i.info) {
			one[j] = i.info[j];
		}
		arr.push(one);
		this.getArr(arr, i.children);
	}
}

var hack = {};
hack.data = [];
hack.amount = 0;
hack.finishAmount = 0;
hack.failURL = [];
hack.realFail=[]

hack.typeInfo = {
	province: "city",
	city: "county",
	county: "town",
	town: "village"
};

hack.start = function () {
	this.init();
	this.fair(hack.data);
	//downloadMoudle.downloadFile('ok.json', JSON.stringify(matas));
}

hack.fair = function (data) {
	let i = 0;
	let fuck=0;
	let oldFail=[];
	let flag=true;
	let time = setInterval(() => {
		if(flag||(hack.failURL.length==0&&hack.finishAmount==hack.amount)){
			if(i>=data.length){
				clearInterval(time);
				console.log("Finish all");
				return;
			}
			console.log("start " + data[i].name);
			hack.amount++;
			hack.getData(data[i]);
			flag=false;
			fuck=0;
			i++;
		}else if(hack.failURL.length!=0&&hack.finishAmount==hack.amount){
			if(fuck==3){
				hack.failURL=[];
				return;
			}
			if(oldFail.length==hack.failURL.length&&oldFail.sort().toString()==hack.failURL.sort().toString()){
				hack.failURL=[];
				return;
			}
			fuck++;
			hack.finishAmount=0;
			oldFail=hack.failURL;
			hack.failURL=[];
			hack.amount=1;
			console.log("restart " + data[i-1].name);
			hack.getData(data[i-1]);
		}
	}, 2000);
}
hack.init = function () {
	let province = document.querySelectorAll('tr.provincetr td');
	for (let i of province) {
		this.data.push(new Mata(i.innerText, "province", {}, i.firstChild.href));
	}
}
hack.getData = function (mata) {
	this.getURL(mata.url, hack.typeInfo[mata.type]).then(function (res) {
		hack.finishAmount++;
		mata.children = res;
		for (let i in res) {
			if (res[i].url)
				// setTimeout(() => {
					hack.getData(res[i]);
				// }, 0 * (i + hack.amount - hack.finishAmount));
			else
				hack.finishAmount++;
		}
		hack.amount += res.length;
		console.log("Finish:" + hack.finishAmount + '/' + hack.amount);
	}, function (info) {
		hack.finishAmount++;
		console.log(info[0]);
		hack.failURL.push(info[1]);
	});

}
hack.getURL = function (url, typeInfo) {
	var xhr = new XMLHttpRequest();
	return new Promise(function (resolve, reject) {
		xhr.open('GET', url, true); //get请求，请求地址，是否异步
		xhr.responseType = "document";

		xhr.onload = function () {
			if (xhr.status == 200) {
				var doc = xhr.response; // 注意:不是oReq.responseText
				if (doc) {
					let data = [];
					let area = doc.querySelectorAll('.' + typeInfo + 'tr td');
					let head = doc.querySelectorAll('.' + typeInfo + 'head td')
					let f = head.length;
					for (let i = 0; i < area.length; i += f) {
						let fm = new Mata(area[i + f - 1].innerText, typeInfo, {}, area[i + f - 1].firstChild.href);
						for (let j = 0; j < (f - 1); j++)
							fm.info[head[j].innerText] = area[i + j].innerText;
						data.push(fm);
					}
					resolve(data);
				} else {
					reject(["parser document faily", url]);
				}
			} else {
				reject(["status!=200", url]);
			}
		};
		xhr.onerror = function (e) {
			reject([e, url]);
		};
		xhr.send();
	});
}