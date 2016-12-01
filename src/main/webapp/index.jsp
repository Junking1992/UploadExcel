<%@ page language="java" contentType="text/html; charset=utf-8"
	pageEncoding="utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta charset="utf-8">
<script src="js/jquery-3.1.1.min.js"></script>
<title>Excel Format</title>
</head>
<body>
	<div id="input" style="text-align: center; margin-top: 200px">
		<h1>上传</h1>
		<input type="file" accept=".xlsx,.xls" name="fileName" id="fileName" /> 
		<input type="button" value="上传坛库存" id="button" onclick="parseExcel()" />
		<input type="button" value="上传坛基本资料" id="button2" onclick="parseExcel2()" />
	</div>
	<div id="msg" style="text-align: center;">
		<p id="now"/>
	</div>
</body>
<script type="text/javascript">

	function parseExcel() {
		var filePath = $("#fileName").val();
		var fileName = filePath.substring(filePath.lastIndexOf('\\') + 1);
		if(fileName == ""){
			alert("请选择Excel!");
			return;
		}
		$.post("change", {
			fileName : fileName,
			action : "A"
		});
		$("#button").hide();
		$("#button2").hide();
		progress();
	}
	
	function parseExcel2() {
		var filePath = $("#fileName").val();
		var fileName = filePath.substring(filePath.lastIndexOf('\\') + 1);
		if(fileName == ""){
			alert("请选择Excel!");
			return;
		}
		$.post("change", {
			fileName : fileName,
			action : "B"
		});
		$("#button").hide();
		$("#button2").hide();
		progress();
	}

	function progress() {
		setTimeout(function() {
			$.get("change", function(data, status) {
				if (status == "success") {
					if(data.startsWith("上传完成")){
						$("#button").show();
						$("#button2").show();
						$("#now").html(data);
					}else if(data.startsWith("错误")){
						$("#button").show();
						$("#button2").show();
						$("#now").html(data);
					}else{
						$("#now").html(data);
						progress();
					}
				}
			});
		}, 1000);
	}

</script>
</html>
