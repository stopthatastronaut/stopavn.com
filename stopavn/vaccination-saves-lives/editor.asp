<%@Language="JScript"%><%
Response.Buffer = true; Response.Expires = -1441;

function init()	{
	// populate the textarea
	var targetPage = Request.Querystring("targetFile").item;
	if(targetPage==null)	{
		return;
	}
	//Response.Write('<p class="title">Source of ' + targetPage + '</p>');
	var targetFile = Server.MapPath(targetPage);
	var objFSO = new ActiveXObject("Scripting.FileSystemObject");
	var objStream = objFSO.opentextFile(targetFile);
	var strPageSource = objStream.readAll()
	objStream.close(); objStream = null;
	objFSO = null;
	Response.Write(Server.HTMLEncode(strPageSource));
}

function saveFile()	{
	var objFSO = new ActiveXObject("Scripting.FileSystemObject");
	// Response.Write(Request("targetFile").item)
	var objStream = objFSO.createTextFile(Server.MapPath(Request("targetFile").item), true);
	objStream.Write(Request.Form("txtFile").item);
	objStream.close();
	objFSO = null;
	Response.Redirect('editor.asp?saved=true');
}

/*
function init()	{
	var targetPage = Request.Querystring("page").item;
	Response.Write('<p class="title">Source of ' + targetPage + '</p>');
	var targetFile = Server.MapPath(targetPage);
	var objFSO = new ActiveXObject("Scripting.FileSystemObject");
	var objStream = objFSO.opentextFile(targetFile);
	var strPageSource = objStream.readAll()
	objStream.close(); objStream = null;
	objFSO = null;
	Response.Write(Server.HTMLEncode(strPageSource));
}
*/

var targetFile = (Request("targetFile").item!=null?Request("targetFile").item:"");

%><html>
<head>
<title>File Editor</title>
<script language="javascript" type="text/javascript">

function goPage(t)	{
	self.location.href = "editor.asp?targetFile=" + t.options[t.selectedIndex].value
}

</script>
</head>
<body>
<% if(Request.QueryString("saved").item=="true") { Response.Write('<p style="color: red; font-weight: bold;">SAVED</p>'); } %>
	<form action="editor.asp?action=saveFile" method="post">
		<p><select name="targetFile" onchange="goPage(this)">
			<option>Select a File</option>
			<option value="/index.html" 
				<%= (targetFile=="/index.html")?'selected="selected"':"" %>>/index.html</option>
			<option value="/vaccination-saves-lives/index.html" 
				<%= (targetFile=="/vaccination-saves-lives/index.html")?'selected="selected"':"" %>>/vsl/index.html</option>
			<option value="/vaccination-saves-lives/woodford-folk-festival.html" 
				<%= (targetFile=="/vaccination-saves-lives/woodford-folk-festival.html")?'selected="selected"':"" %>>/vsl/WFF.html</option>
			<option value="/vaccination-saves-lives/bloggers.html" 
				<%= (targetFile=="/vaccination-saves-lives/bloggers.html")?'selected="selected"':"" %>>/vsl/bloogers.html</option>
			<option value="/style.css" 
				<%= (targetFile=="/style.css")?'selected="selected"':"" %>>/style.css</option>
			<option value="/robots.txt" 
				<%= (targetFile=="/robots.txt")?'selected="selected"':"" %>>/robots.txt</option>
			<option value="/testfile.txt" 
				<%= (targetFile=="/testfile.txt")?'selected="selected"':"" %>>/testfile.txt</option>
		</select></p>
		<textarea name="txtFile" style="width: 500px; height: 350px;"><% 
			var action = new String(Request.Querystring("action"));				
			(action!='undefined'&&action.length<20)?eval(action + '()'):init();	
		%></textarea><br />
		<input type="submit" value="SAVE" />
	</form>
</body>
</html>