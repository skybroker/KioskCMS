function GetAllChecked() {
    var retstr = "";
    var tb = document.getElementById("GridView1");
    var j = 0;
    for (var i = 0; i < tb.rows.length; i++) {
        var objtr = tb.rows[i];
        if (objtr.cells.length < 1)
            continue;
		for (var l=0;l<objtr.cells.length;l++)
		{
			var objtd = objtr.cells[l];
			for (var k = 0; k < objtd.childNodes.length; k++) {
				var objnd = objtd.childNodes[k];
				if (objnd.type == "checkbox") {
					if (objnd.checked) {
						if (j > 0)
							retstr += ",";
						retstr += objnd.value;
						j++;
					}
					break;
				}
			}
		}
    }
    return retstr;
}
function doCheckAll(obj) {
    var form = obj.form;
    for (var i = 0; i < form.elements.length; i++) {
        var e = form.elements[i];
        e.checked = obj.checked;
    }
}
function OpenImageBrowser(Element)
{
	if (navigator.appName=="Microsoft Internet Explorer")
	{
		window.showModalDialog("Sys_FileManager.asp?Element="+Element,window,"dialogWidth:600px;dialogHeight:330px;center:yes")
	}
	else
	{
		window.open("Sys_FileManager.asp?Element="+Element,"","toolbar=no,location=no,directories=no,status=yes,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=600,height=330;center=yes");
	}
}
/*function FCKeditor_OnComplete(editorInstance)
{
	editorInstance.Events.AttachEvent('OnBlur',FCKeditor_OnBlur);
	editorInstance.Events.AttachEvent('OnFocus',FCKeditor_OnFocus);
}
function FCKeditor_OnBlur(editorInstance)
{
	editorInstance.ToolbarSet.Collapse();
}
function FCKeditor_OnFocus(editorInstance)
{
	editorInstance.ToolbarSet.Expand();
}*/