<script language="javascript" runat="server">
/**
* jsonencode.asp
*
* JavaScript function for encoding strings for JSON
* using fast JavaScript callbacks on a single regular expression search/replace per string
*
* @version 1.00.01 2011-03-16
* @package asp
* @author Ross McKay <rmckay@webaware.com.au>
* @link https://github.com/webaware/ASP-JSONEncode
* @copyright copyright © 2011 WebAware Pty Ltd
*
* This library is free software; you can redistribute it and/or
* modify it under the terms of the GNU Lesser General Public
* License as published by the Free Software Foundation; either
* version 2.1 of the License, or (at your option) any later version.
*
* This library is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
* Lesser General Public License for more details.
*
* You should have received a copy of the GNU Lesser General Public
* License along with this library; if not, write to the Free Software
* Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
*
* Full license: {@link http://www.webaware.com.au/free/license.htm}
*/

/**
* encode a string for embedding as a value in a JSON document
*
* will encode control characters as Unicode hex, and LF CR TAB FF BS / \ " as readable escaped characters
*
* ref: http://www.ietf.org/rfc/rfc4627.txt?number=4627
*
* @return string
* @param string src the value to encode
*/
function JSONEncodeString(src) {
	if (null === src || typeof src == "undefined")
		return "";

	var s = String(src);

	// encode control characters as well as \ / "
	return s.replace(/[\\\/"\x00-\x1f\x7f-\xa0\u2000-\u200f\u2028-\u202f]/g, function(match) {
		switch (match) {
			case "\\":
				return "\\\\";
			case "/":
				return "\\/";
			case '"':
				return '\\"';
			case "\r":
				return "\\r";
			case "\n":
				return "\\n";
			case "\t":
				return "\\t";
			case "\f":
				return "\\f";
			case "\b":
				return "\\b";
			default:
				// return as \uNNNN
				var c = match.charCodeAt(0);
				return "\\u" + ("0000" + c.toString(16)).slice(-4);
		}
	});
}

</script>

<script language="vbscript" runat="server">

'-----------------------------------------------------------------------
' format a date object into ISO-8601 so that JavaScript will parse it
'
' @param Date d
' @return string
'-----------------------------------------------------------------------
Function JSONEncodeDate(d)
	JSONEncodeDate = Right("000" & Year(d), 4) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2) _
		& "T" & Right("0" & Hour(d), 2) & "-" & Right("0" & Minute(d), 2) & "-" & Right("0" & Second(d), 2)
End Function

'-----------------------------------------------------------------------
' convert a dictionary object into a collection of JSON elements
'
' @param string elementName the name of the JSON element wrapping the collection
' @param Scripting.Dictionary dict a name/value pair collection of mixed data types
' @return string
'-----------------------------------------------------------------------
Function JSONEncodeDict(ByVal elementName, ByVal dict)
	Dim i, delim

	JSONEncodeDict = """" & JSONEncodeString(elementName) & """:{"
	delim = ""
	For Each i In dict
		JSONEncodeDict = JSONEncodeDict & delim
		Select Case VarType(dict(i))
		Case vbObject
			JSONEncodeDict = JSONEncodeDict & JSONEncodeDict(i, dict(i))
		Case vbNull
			JSONEncodeDict = JSONEncodeDict & """" & JSONEncodeString(i) & """:null"
		Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbByte
			JSONEncodeDict = JSONEncodeDict & """" & JSONEncodeString(i) & """:" & dict(i)
		Case vbDate
			JSONEncodeDict = JSONEncodeDict & """" & JSONEncodeString(i) & """:" & JSONEncodeDate(dict(i))
		Case vbBoolean
			JSONEncodeDict = JSONEncodeDict & """" & JSONEncodeString(i) & """:" & LCase(dict(i))
		Case Else
			If IsArray(dict(i)) Then
				JSONEncodeDict = JSONEncodeDict & """" & JSONEncodeString(i) & """:" & JSONEncodeArray(dict(i))
			Else
				JSONEncodeDict = JSONEncodeDict & """" & JSONEncodeString(i) & """:""" & JSONEncodeString(dict(i)) & """"
			End If
		End Select
		delim = ","
	Next
	JSONEncodeDict = JSONEncodeDict & "}"
End Function

'-----------------------------------------------------------------------
' convert an array into a collection of JSON elements,
'
' @param Array arr an array of mixed data types
' @return string
'-----------------------------------------------------------------------
Function JSONEncodeArray(ByVal arr)
	Dim i, delim

	JSONEncodeArray = "["
	delim = ""
	For i = LBound(arr) To UBound(arr)
		JSONEncodeArray = JSONEncodeArray & delim

		Select Case VarType(arr(i))
		Case vbObject
			JSONEncodeArray = JSONEncodeArray & JSONEncodeDict(i, arr(i))
		Case vbNull
			JSONEncodeArray = JSONEncodeArray & "null"
		Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbByte
			JSONEncodeArray = JSONEncodeArray & arr(i)
		Case vbDate
			JSONEncodeArray = JSONEncodeArray & JSONEncodeDate(arr(i))
		Case vbBoolean
			JSONEncodeArray = JSONEncodeArray & LCase(arr(i))
		Case Else
			If IsArray(arr(i)) Then
				JSONEncodeArray = JSONEncodeArray & JSONEncodeArray(arr(i))
			Else
				JSONEncodeArray = JSONEncodeArray & """" & JSONEncodeString(arr(i)) & """"
			End If
		End Select

		delim = ","
	Next
	JSONEncodeArray = JSONEncodeArray & "]"
End Function

</script>
