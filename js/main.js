var area; //使用面积
var cool_per_area; //单位制冷量
var cool_load; //设计冷负荷
var shu; //需求台数
var worksheet;
var datas = []; //磁悬浮数据
var datas2 = []; //离心机数据
var datas_mag = []; //磁悬浮选型数据
var datas_cen = []; //离心机选型数据
var iplv_total_maglev = 0; //磁悬浮=设计制冷负荷/IPLV
var iplv_total_centrifugenull = 0; //离心机=设计制冷负荷/IPLV
var total_power1 = 0,
	total_power2 = 0; //总电量
var invisible_cost = 0; //其他隐形费用节省
var max_mag = 5276; //磁悬浮最大制冷量
var max_cen = 9850; //离心机最大制冷量
var type_area = [
	[90, 120, 120, 120, 140],
	[150, 230, 230, 200, 250],
	[600, 700, 700, 700, 800],
	[100, 120, 120, 110, 140],
	[80, 100, 100, 100, 120]
];
$(document).ready(function() {
	(function($) {
		$.dataExcel = function() {};
		$.extend($.dataExcel, {
			name: 5
		}, {
			cool: 5
		}, {
			iplv: 5
		}, {
			seq: 5
		}, {
			desp: 5
		}, {
			imgsrc: 5
		});
	})(jQuery);
	/* set up XMLHttpRequest */
	var url = "db/db.xlsx";
	var url2 = "db/db2.xlsx";
	var oReq = new XMLHttpRequest();
	oReq.open("GET", url, true);
	oReq.responseType = "arraybuffer";
	generateDatas(url, 1);
	generateDatas(url2, 2);
	//点击磁悬浮
	$("#maglev").click(function() {
		if (!valideForm()) {
			return false;
		}
		xuanxing();
		var total = area * cool_per_area / 1000;
		var avg = total / shu;
		if (avg > max_mag) {
			var sug_shu = total / max_mag;
			alert("平均为制冷量：" + avg.toFixed(2) + "，超过磁悬浮最大制冷量：" + max_mag + "，建议 " + Math.ceil(sug_shu).toFixed(0) + " 台。");
			datas_mag = [];
			$("#shu").focus();
			return false;
		}
		$("#img-xuanxing").attr("src", "./img/meglev_big_icon.png");
		//				location.hash = "pageTwo";
		var ul = $("#pagetwo-ul");
		//清除<li>
		ul.empty();
		$(".ui-header").addClass("ui-fixed-hidden");
		for (var i = 0; i < shu; i++) {
			var data = datas_mag[i];
			var seq = data.seq;
			var li = $("<li style=\"height:120px;\"/>", {
				"data-icon": false
					//				"class": "ui-li-static ui-body-inherit ui-first-child"
			});
			var div = $("<div class=\"ui-field-contain ui-main-list div-select\"/> ").appendTo(li);
			var div_left = $("<div class=\"div-left\"/>").appendTo(div);
			var div_right = $("<div class=\"div-right\" />").appendTo(div);
			var img = $("<img src=\"img/meglev_small_icon.png\" width=\"50px\"/>").appendTo(div_left);
			var p1 = $("<div class=\"select-div\"><span style=\"float:left;font-weight:bold;margin-top:10px;\">磁悬浮选型：</span></div>").appendTo(div_right);
			var selectH = $("<select class=\"select_xh\" uniq=\"select" + i + "\" data-native-menu=\"false\"/> ").appendTo(p1);
			//			if (seq > 1) {
			//				var option = $("<option value=" + (seq - 2) + ">" + datas[seq - 2].name + "</option>").appendTo(selectH);
			//			}
			//			if (seq > 0) {
			//				var option = $("<option value=" + (seq - 1) + ">" + datas[seq - 1].name + "</option>").appendTo(selectH);
			//			}
			//			var option = $("<option value=" + seq + " selected=\"true\">" + data.name + "</option>").appendTo(selectH);
			//			if (seq < datas.length - 1) {
			//				var option = $("<option value=" + (seq + 1) + ">" + datas[seq + 1].name + "</option>").appendTo(selectH);
			//			}
			//			if (seq < datas.length - 2) {
			//				var option = $("<option value=" + (seq + 2) + ">" + datas[seq + 2].name + "</option>").appendTo(selectH);
			//			}
			for (var j = 0; j < datas.length; j++) {
				if (j == seq) {
					$("<option value=" + j + " selected=\"true\">" + data.name + "</option>").appendTo(selectH);
				} else {
					$("<option value=" + j + ">" + datas[j].name + "</option>").appendTo(selectH);
				}
			}
			var p0 = $("<p ><span style=\"font-weight:bold;\" id=\"xh-desp\">" + data.desp + "型号：</span><span class=\"select_xinhao\" id=\"xinghao" + i + "\">" + data.name + "</span></p>").appendTo(div_right);
			var p2 = $("<p ><span style=\"font-weight:bold;\">制冷量：</span><span class=\"select_cool\" id=\"cool" + i + "\">" + data.cool + "</span> KW</p>").appendTo(div_right);
			var p3 = $("<p><span style=\"font-weight:bold;\">IPLV：</span><span class=\"select_iplv\" id=\"iplv" + i + "\">" + data.iplv + "</span></p>").appendTo(div_right);
			li.appendTo(ul);

		}
		//		$("#pagetwo-ul").listview("refresh");.trigger( "create" );
		//		$.mobile.changePage("#pagetwo-ul");
		$("#pagetwo-ul").listview("refresh");
		$(".ui-field-contain").trigger("create");
	});
	//点击离心机
	$("#centrifuge").click(function() {
		if (!valideForm()) {
			return false;
		}
		xuanxing();
		var total = area * cool_per_area / 1000;
		var avg = total / shu;
		if (avg > max_cen) {
			var sug_shu = total / max_cen;
			alert("平均为制冷量：" + avg.toFixed(2) + "，超过离心机最大制冷量：" + max_cen + "，建议 " + Math.ceil(sug_shu).toFixed(0) + " 台。");
			datas_cen = [];
			$("#shu").focus();
			return false;
		}
		$("#img-xuanxing").attr("src", "./img/lx.png");
		//		location.hash = "pageTwo";
		var ul = $("#pagetwo-ul");
		//清除<li>
		ul.empty();
		$(".ui-header").addClass("ui-fixed-hidden");
		for (var i = 0; i < shu; i++) {
			var data = datas_cen[i];
			var seq = data.seq;
			var li = $("<li style=\"height:120px;\"/>", {
				"data-icon": false
					//				"class": "ui-li-static ui-body-inherit ui-first-child"
			});

			var div = $("<div class=\"ui-field-contain ui-main-list div-select\"/> ").appendTo(li);
			var div_left = $("<div class=\"div-left\"/>").appendTo(div);
			var div_right = $("<div class=\"div-right\" />").appendTo(div);
			var img = $("<img src=\"" + data.imgsrc + "\" width=\"50px\" id=\"imgsrc2-" + i + "\"/> ").appendTo(div_left);

			var p1 = $("<div class=\"select-div\"><span style=\"float:left;font-weight:bold;margin-top:10px;\">普通机选型：</span></div>").appendTo(div_right);
			var selectH = $("<select class=\"select2_xh\" uniq=\"select2" + i + "\" data-native-menu=\"false\"/> ").appendTo(p1);
			//			if (seq > 1) {
			//				var option = $("<option value=" + (seq - 2) + ">" + datas2[seq - 2].name + "</option>").appendTo(selectH);
			//			}
			//			if (seq > 0) {
			//				var option = $("<option value=" + (seq - 1) + ">" + datas2[seq - 1].name + "</option>").appendTo(selectH);
			//			}
			//			var option = $("<option value=" + seq + " selected=\"true\">" + data.name + "</option>").appendTo(selectH);
			//			if (seq < datas2.length - 1) {
			//				var option = $("<option value=" + (seq + 1) + ">" + datas2[seq + 1].name + "</option>").appendTo(selectH);
			//			}
			//			if (seq < datas2.length - 2) {
			//				var option = $("<option value=" + (seq + 2) + ">" + datas2[seq + 2].name + "</option>").appendTo(selectH);
			//			}
			for (var j = 0; j < datas2.length; j++) {
				if (j == seq) {
					$("<option value=" + j + " selected=\"true\">" + data.name + "</option>").appendTo(selectH);
				} else {
					$("<option value=" + j + ">" + datas2[j].name + "</option>").appendTo(selectH);
				}
			}
			var p0 = $("<p><span style=\"font-weight:bold;\" id=\"xh-desp2-" + i + "\">" + data.desp + "型号：</span><span class=\"select_xinhao\" id=\"xinghao2" + i + "\">" + data.name + "</span></p>").appendTo(div_right);
			var p2 = $("<p><span style=\"font-weight:bold;\">制冷量：</span><span class=\"select_cool\" id=\"cool2" + i + "\">" + data.cool + "</span> KW</p>").appendTo(div_right);
			var p3 = $("<p><span style=\"font-weight:bold;\">IPLV：</span><span class=\"select_iplv\" id=\"iplv2" + i + "\">" + data.iplv + "</span></p>").appendTo(div_right);
			li.appendTo(ul);
		}
		//		$("#pagetwo-ul").listview("refresh");
		$("#pagetwo-ul").listview("refresh");
		$(".ui-field-contain").trigger("create");
	});
	$(".budget-a").click(function() {
		//		location.hash = "pageThree";
		if (datas_mag.length <= 0 || datas_cen.length <= 0) {
			alert("请先进行磁悬浮选型。");
			return false;
		}
		calculatePower();

	});
	get_unit_cool();
	changeArea();
	//选择建筑类型，地址区域
	$("#build-select,#addr-select").change(function() {
		get_unit_cool();
		changeArea();
	});
	//修改面积
	$("#area").bind('input propertychange', function() {
		changeArea();
	});
	//修改单位制冷量
	$("#unit-cool").bind('input propertychange', function() {
		changeArea();
	});
	$(".calculate-item").bind('input propertychange', function() {
		calculatePower();
	});
	calculatePower();
	$(".calculate-item").bind('blur', function() {
		var daysinyear = $("#daysinyear").val().trim();
		var hour = $("#hour").val().trim();
		if (daysinyear > 366) {
			alert("年运行天数不能大于366天。");
			$("#daysinyear").focus();
			$("#daysinyear").val("366");
			return;
		}
		if (hour > 24) {
			alert("每天运行小时数不能大于24小时。");
			$("#hour").focus();
			$("#hour").val("24");
			return;
		}
	});
	//确定预算，计算回收期
	$("#submit_btn").click(function() {
		//购买预算
		var yusuan1 = $("#yusuan1").val().trim();
		if (!isValidValue(yusuan1)) {
			alert("请输入正确的预算（大于0的数字）。");
			$("#yusuan1").focus();
			return;
		}
		//购买预算
		var yusuan2 = $("#yusuan2").val().trim();
		if (!isValidValue(yusuan2)) {
			alert("请输入正确的预算（大于0的数字）。");
			$("#yusuan2").focus();
			return;
		}
		//每年电费节省
		var power_save = total_power2 - total_power1;
		//维保节省
		var weibao1 = $("#weibao1").val().trim();
		var weibao2 = $("#weibao2").val().trim();
		if (!isValidValue(weibao1)) {
			alert("请输入正确的维保费用（大于0的数字）。");
			$("#weibao1").focus();
			return;
		}
		if (!isValidValue(weibao2)) {
			alert("请输入正确的维保费用（大于0的数字）。");
			$("#weibao2").focus();
			return;
		}
		var weibao_save = weibao2 - weibao1;
		//年节约费用
		var save_per_year = power_save + weibao_save + invisible_cost;
		var payback = (yusuan1 - yusuan2) / save_per_year;
		$("#budget-result").html(payback.toFixed(2) + " 年");
		var h = $(document).height() - $(window).height();
		$(document).scrollTop(h);
	});
	//磁悬浮自定义型号
	$(document).on('change', '.select_xh', function(e) {
		var id = $(this).attr("uniq");
		var i = id.substr(6, 1);
		var value = $(this).val();
		$("#xinghao" + i).html(datas[value].name);
		$("#cool" + i).html(datas[value].cool);
		$("#iplv" + i).html(datas[value].iplv);
		var data = new $.dataExcel();
		datas_mag[i] = datas[value];
		iplv_total_maglev = 0;
		//重新计算iplv
		for (var j = 0; j < datas_mag.length; j++) {
			iplv_total_maglev += datas_mag[j].cool / datas_mag[j].iplv;
		}
	});
	//离心机自定义型号
	$(document).on('change', '.select2_xh', function(e) {
		var id = $(this).attr("uniq");
		var i = id.substr(7, 1);
		var value = $(this).val();
		$("#xinghao2" + i).html(datas2[value].name);
		$("#cool2" + i).html(datas2[value].cool);
		$("#iplv2" + i).html(datas2[value].iplv);
		$("#xh-desp2-" + i).html(datas2[value].desp + "型号：");
		$("#imgsrc2-" + i).attr("src", datas2[value].imgsrc);
		var data = new $.dataExcel();
		datas_cen[i] = datas2[value];
		iplv_total_centrifugenull = 0;
		//重新计算iplv
		for (var j = 0; j < datas_cen.length; j++) {
			iplv_total_centrifugenull += datas_cen[j].cool / datas_cen[j].iplv;
		}
	});
	//

});
//磁悬浮或者离心机选型
function xuanxing() {
	iplv_total_maglev = 0; //磁悬浮
	iplv_total_centrifugenull = 0; //离心机
	var total = area * cool_per_area / 1000;
	var total_temp = total;
	var shu_temp = shu;
	var avg = total / shu;
	var j = 0;
	if (avg > max_mag) {
		datas_mag = [];
	} else {
		for (var i = 0; i < shu; i++) {
			var ii1 = queryM1(total_temp / shu_temp);
			var data = datas[ii1];
			total_temp = total_temp - data.cool;
			shu_temp--;
			datas_mag[j] = data;
			iplv_total_maglev += data.cool / data.iplv;
			j++;
		}
	}

	var total_temp = total;
	var shu_temp = shu;
	var j = 0;
	if (avg > max_cen) {
		datas_cen = [];
	} else {
		for (var i = 0; i < shu; i++) {
			var ii2 = queryM2(total_temp / shu_temp);
			var data = datas2[ii2];
			total_temp = total_temp - data.cool;
			shu_temp--;
			datas_cen[j] = data;
			iplv_total_centrifugenull += data.cool / data.iplv;
			j++;
		}
	}

}
//当数据发生改变时，重新计算总耗电量
function calculatePower() {
	var daysinyear = $("#daysinyear").val().trim();
	var hour = $("#hour").val().trim();
	var price = $("#price").val().trim();
	var xishu = $("#xishu").val().trim();
	var yusuan1 = $("#yusuan1").val().trim();
	var yusuan2 = $("#yusuan2").val().trim();
	var $total_kwh1 = $("#total_kwh1");
	var $total_kwh2 = $("#total_kwh2");
	var $total_rmb1 = $("#total_rmb1");
	var $total_rmb2 = $("#total_rmb2");
	if (isValidValue(daysinyear) && isValidValue(hour) && isValidValue(price) && isValidValue(xishu) && isValidValue(iplv_total_maglev) && isValidValue(iplv_total_centrifugenull)) {
		var total_time = daysinyear * hour;
		total_power1 = iplv_total_maglev * xishu * total_time / 10000;
		total_power2 = iplv_total_centrifugenull * xishu * total_time / 10000;
		var rmb1 = total_power1 * price;
		var rmb2 = total_power2 * price;
		$total_kwh1.html(total_power1.toFixed(2));
		$total_kwh2.html(total_power2.toFixed(2));
		$total_rmb1.html(rmb1.toFixed(2));
		$total_rmb2.html(rmb2.toFixed(2));
	} else {
		$total_kwh1.html(0);
		$total_kwh2.html(0);
		$total_rmb1.html(0);
		$total_rmb2.html(0);
	}
}
//改变面积或者单位制冷量
function changeArea() {
	area = $("#area").val().trim();
	cool_per_area = $("#unit-cool").val().trim();
	if ((area != "" && !isNaN(area) && area > 0) &&
		(cool_per_area != "" && !isNaN(cool_per_area) && cool_per_area > 0)) {
		$("#cool-load").val(area * cool_per_area / 1000);
		$("#cool-load-show").html(area * cool_per_area / 1000);
	} else {
		$("#cool-load").val(0);
		$("#cool-load-show").html(0);
	}
}

function isValidValue(value) {
	return value != "" && !isNaN(value) && value > 0;
}
//检查表单的合法性
function valideForm() {
	area = $("#area").val().trim();

	if (area == "") {
		alert("请输入使用面积。");
		$("#area").focus();
		return false;
	} else if (isNaN(area) || area <= 0) {
		alert("使用面积必须是大于0的数字！");
		$("#area").focus();
		return false;
	}
	cool_per_area = $("#unit-cool").val().trim();
	if (cool_per_area == "") {
		alert("单位请输入制冷量。");
		$("#unit-cool").focus();
		return false;
	} else if (isNaN(cool_per_area) || cool_per_area <= 0) {
		alert("单位制冷量必须是大于0的数字！");
		$("#unit-cool").focus();
		return false;
	}
	cool_load = $("#cool-load").val().trim();
	//	if (cool_load == "") {
	//		alert("设计冷负荷不能为空！");
	//		$("#cool-load").focus();
	//		return false;
	//	} else if (isNaN(cool_load) || cool_load <= 0) {
	//		alert("设计冷负荷必须是大于0的数字！");
	//		$("#cool-load").focus();
	//		return false;
	//	}
	shu = $("#shu").val().trim();
	if (shu == "") {
		alert("请输入需求台数。");
		$("#shu").focus();
		return false;
	} else if (isNaN(shu) || shu <= 0) {
		alert("需求台数必须是大于0的数字！");
		$("#shu").focus();
		return false;
	}
	return true;
}
//选择磁悬浮
function queryM1(output) {
	for (var i = 0; i < datas.length; i++) {
		if (output <= datas[i].cool) {
			return i;
		}
	}
	return 0;
}
//选择离心机
function queryM2(output) {
	for (var i = 0; i < datas2.length; i++) {
		if (output <= datas2[i].cool) {
			return i;
		}
	}
	return 0;
}
//读取excel
function generateDatas(url, flag) {
	var oReq = new XMLHttpRequest();
	oReq.open("GET", url, true);
	oReq.responseType = "arraybuffer";
	oReq.onload = function(e) {
		var arraybuffer = oReq.response;
		/* convert data to binary string */
		var data = new Uint8Array(arraybuffer);
		var arr = new Array();
		for (var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
		var bstr = arr.join("");
		/* Call XLSX */
		var workbook = XLSX.read(bstr, {
			type: "binary"
		});
		/* DO SOMETHING WITH workbook HERE */
		var first_sheet_name = workbook.SheetNames[0];
		var address_of_cell = 'AA1';
		/* Get worksheet */
		worksheet = workbook.Sheets[first_sheet_name];
		/* Find desired  cell */
		var desired_cell = worksheet[address_of_cell];
		/* Get the value */
		var desired_value = desired_cell.v;
		//		alert("value:" + desired_value);
		var j = 0;
		for (var i = 69; i < 196; i++) {
			var ch = null;
			if (i < 91) {
				ch = String.fromCharCode(i);
			} else {
				ch = 'A' + String.fromCharCode(i - 26);
			}
			desired_cell_1 = ch + '1';
			var desired_cell_2 = ch + '2';
			var desired_cell_5 = ch + '5';
			var value1 = worksheet[desired_cell_1];
			var value2 = worksheet[desired_cell_2];
			var value5 = worksheet[desired_cell_5];
			if (typeof(value1) == "undefined") {
				break;
			}
			var data = new $.dataExcel();
			data.cool = value2.v;
			data.iplv = value5.v;
			data.seq = j;
			if (flag == 1) {
				//型号LSBLX*/R4(BP)
				data.name = "LSBLX" + value1.v + "/R4(BP)";
				data.desp = "磁悬浮";
				datas[j] = data;
			} else {
				if (data.cool < 1760) {
					//LSBLG*/R4G
					data.name = "LSBLG" + value1.v + "/R4G";
					data.desp = "螺杆机";
					data.imgsrc = "img/lg.png";
				} else {
					//LSBLX-*S/R4A
					data.name = "LSBLX-" + value1.v + "S/R4A";
					data.desp = "离心机";
					data.imgsrc = "img/lx.png";
				}
				datas2[j] = data;
			}
			j++;
		}
	}
	oReq.send();
}
//获得单位制冷量
function get_unit_cool() {
	var type = $("#build-select").val();
	var addr = $("#addr-select").val();
	var $unit_cool = $("#unit-cool");
	var cool_value = 0;
	cool_value = type_area[type][addr];
	$unit_cool.val(cool_value);
}