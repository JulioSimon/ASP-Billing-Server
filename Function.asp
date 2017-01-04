<%
Function FilterReqXSS(str)

str = Replace(str, "<", "")
str = Replace(str, ">", "")
str = Replace(str, "(", "")
str = Replace(str, ")", "")
str = Replace(str, "{", "")
str = Replace(str, "}", "")
str = Replace(str, "'", "")
str = Replace(str, "%", "")
str = Replace(str, """", "")
str = Replace(str, "!", "")
str = Replace(str, "+", "")
str = Replace(str, ":", "")
str = Replace(str, ";", "")
str = Replace(str, "=", "")
str = Replace(str, "&", "")

FilterReqXSS = str

End Function



Function FilterReqXSS2(str)

str = Replace(str, "<", "")
str = Replace(str, ">", "")
str = Replace(str, "(", "")
str = Replace(str, ")", "")
str = Replace(str, "{", "")
str = Replace(str, "}", "")
str = Replace(str, "'", "")
str = Replace(str, "%", "")
str = Replace(str, """", "")
str = Replace(str, "!", "")
str = Replace(str, "+", "")
str = Replace(str, ";", "")

FilterReqXSS2 = str

End Function



Function FilterSQL(str)

FilterSQL = Replace(str, "'", "''")

End Function



Function FilterSQL_DQ(str)

FilterSQL_DQ = Replace(str, "'", "''''")

End Function



Function FilterHtml(str)

str = Replace(str, "&", "&amp;")
str = Replace(str, "%", "&#37")
str = Replace(str,"<","&lt;")
str = Replace(str,">","&gt;")
str = Replace(str, chr(39), "&#39")
str = Replace(str, chr(34), "&#34")
str = Replace(str, vbcrlf, "<br>" )

FilterHtml = str

End Function



Function FilterHtmlPer(str)

str = Replace(str, vbcrlf, "<br>" )
str = Replace(str,"  ","&nbsp;&nbsp;")

FilterHtmlPer = str

End Function



Function FilterScript(str)

str = Replace(str,"<script","<x-script")
str = Replace(str,"</script","</x-script")
str = Replace(str, vbcrlf, "<br>" )

FilterScript = str

End Function
%>