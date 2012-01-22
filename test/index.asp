<%@ language="vbscript" codepage=65001 %>
<%
Option Explicit
Response.CharSet = "utf-8"
%>
<!-- #include file="../jsonencode.asp" -->
<%
Class ThisPage
	Private m_strings
	Private m_mixed
	Private m_dict
	Private m_dictNested

	'----------------------------------------------------------------------
	Private Sub Class_Initialize()
		' an array with example strings
		m_strings = Array("Albatros", """Fingers"" O'Malley", "tab " & vbTab & " lf " & vbLf & " cr " & vbCr, "\/")

		' an array with mixed data types
		m_mixed = Array("yellow", 12.3, Now, Null)

		' a simple dictionary
		Set m_dict = Server.CreateObject("Scripting.Dictionary")
		m_dict.Add "a", "albatros"
		m_dict.Add "b", "buffalo"
		m_dict.Add "c", "civet"
		m_dict.Add "d", "donkey"
		m_dict.Add "none", Null

		' a nested dictionary
		Set m_dictNested = Server.CreateObject("Scripting.Dictionary")
		m_dictNested.Add "strings", m_strings
		m_dictNested.Add "mixed", m_mixed
		m_dictNested.Add "animals", m_dict
	End Sub

	'----------------------------------------------------------------------
	Public Sub showStrings()
		Dim i

		For Each i In m_strings
			Response.Write Server.HTMLEncode(JSONEncodeString(i)) & vbCrLf
		Next
	End Sub

	'----------------------------------------------------------------------
	Public Sub showMixed()
		Dim i

		For Each i In m_mixed
			Response.Write Server.HTMLEncode(JSONEncodeString(i)) & vbCrLf
		Next
	End Sub

	'----------------------------------------------------------------------
	Public Sub showArray()
		Response.Write Server.HTMLEncode(JSONEncodeArray(m_strings)) & vbCrLf
	End Sub

	'----------------------------------------------------------------------
	Public Sub showDict()
		Response.Write Server.HTMLEncode(JSONEncodeDict("simple", m_dict)) & vbCrLf
	End Sub

	'----------------------------------------------------------------------
	Public Sub showDictNested()
		Response.Write Server.HTMLEncode(JSONEncodeDict("nested", m_dictNested)) & vbCrLf
	End Sub

End Class

Dim p : Set p = new ThisPage
%>
<!DOCTYPE html>
<html lang="en-au">

<head>
<title>Test JSON Encoding</title>
<meta charset="utf-8" />
<link rel="stylesheet" href="simple.css" />
</head>

<body>

<h2>Strings</h2>
<pre>
<% p.showStrings() %>
</pre>

<h2>Mixed Types</h2>
<pre>
<% p.showMixed() %>
</pre>

<h2>Array</h2>
<pre>
<% p.showArray() %>
</pre>

<h2>Dictionary</h2>
<pre>
<% p.showDict() %>
</pre>

<h2>Nested Dictionary</h2>
<pre>
<% p.showDictNested() %>
</pre>

</body>
</html>
