<#include "/includes/header.ftl" />
	<script language="javascript">
		var btns = new Array();
		var btIndex = 0;
		btns[btIndex++] = new btnItem('上传','mysave()','btn_save');	
		AddHeader('Excel导入',btns);
	
		function mysave() {
			var fileName = document.forms[0].excelFile.value;
			if(fileName.length < 1 || !fileName.match(/^(.*)(\.)(.{1,8})$/)){
				alert('请选择一个Excel文件！');
				return;
			}
			if(fileName.substring(fileName.lastIndexOf('.')+1).toLowerCase() != 'xls'){
				alert('请选择一个Excel文件！');
				return;
			}
			
			document.forms[0].excelFile.disable = true;
			if(document.forms[0].checkForm()){
				document.forms[0].submit();
			}
		}
	</script>  
	<span class="right_middle">
	<br>
	<@ww.form name="'myform'" namespace="'/sys4'" action="'ExcelImport'" method="'POST'" enctype="'multipart/form-data'">
		<@ww.hidden name="'sheetIndex'" />
		<@ww.hidden name="'xmlConfig'" />
		<@ww.hidden name="'redirect'" />
		<table cellspacing="0" cellpadding="3" class="tform" width="90%" align="center">
			<thead>
				<tr>
					<td align="center" width="25%">项</td>
					<td align="center" width="75%">内容</td>
				</tr>
			</thead>
			<tbody>
				<tr>
					<td align="right" nowrap>选择文件</td>
					<td><@ww.file name="'excelFile'" size="35" accept="'*.xls'"/></td>
				</tr>
			</tbody>
		</table>
	</@ww.form>
	</span>
	<#include "/includes/error.ftl" />
	<script language="javascript">
		AddFooter('');
	</script>
<#include "/includes/footer.ftl" />