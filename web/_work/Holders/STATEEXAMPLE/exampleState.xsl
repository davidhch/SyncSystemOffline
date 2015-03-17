<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:output method="html" encoding="iso-8859-1" doctype-public="-//W3C//DTD XHTML 1.0 Transitional//EN" doctype-system="http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"/>
<xsl:template match="/">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>
<title>new subscriber</title>
<!--//note JS may be moved elsewhere, but for now it is here//-->
<script language="JavaScript" type="text/javascript" src="JSGoesHere.js"></script>
<style>#navarea{color:#009900;}</style>
</head>
<body>
nav state loaded
<form name="FormName" method="post" action="example.asp">

	<input>
		<xsl:attribute name="ID"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/save/id"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/save/name"/></xsl:attribute>
		<xsl:attribute name="TYPE"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/save/type"/></xsl:attribute>
	</input>
	
	<input>
		<xsl:attribute name="ID"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/email-address/id"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/email-address/name"/></xsl:attribute>
		<xsl:attribute name="TYPE"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/email-address/type"/></xsl:attribute>
		<xsl:attribute name="VALUE"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/email-address/value"/></xsl:attribute>
	</input>
	
	<input>
		<xsl:attribute name="ID"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/mobile-phone/id"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/mobile-phone/name"/></xsl:attribute>
		<xsl:attribute name="TYPE"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/mobile-phone/type"/></xsl:attribute>
		<xsl:attribute name="VALUE"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/mobile-phone/value"/></xsl:attribute>
	</input>
	
	<input>
		<xsl:attribute name="ID"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/receive-email/id"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/receive-email/name"/></xsl:attribute>
		<xsl:attribute name="TYPE"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/receive-email/type"/></xsl:attribute>
		<xsl:attribute name="VALUE"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/receive-email/value"/></xsl:attribute>
	</input>
	
	<input>
		<xsl:attribute name="ID"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/title/id"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/title/name"/></xsl:attribute>
		<xsl:attribute name="TYPE"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/title/type"/></xsl:attribute>
		<xsl:attribute name="VALUE"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/title/value"/></xsl:attribute>
	</input>
	<input>
		<xsl:attribute name="ID"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/first-name/id"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/first-name/name"/></xsl:attribute>
		<xsl:attribute name="TYPE"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/first-name/type"/></xsl:attribute>
		<xsl:attribute name="VALUE"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/first-name/value"/></xsl:attribute>
	</input>	
	
	<input>
		<xsl:attribute name="ID"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/last-name/id"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/last-name/name"/></xsl:attribute>
		<xsl:attribute name="TYPE"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/last-name/type"/></xsl:attribute>
		<xsl:attribute name="VALUE"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/last-name/value"/></xsl:attribute>
	</input>	
	
	<input>
		<xsl:attribute name="ID"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/company-name/id"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/company-name/name"/></xsl:attribute>
		<xsl:attribute name="TYPE"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/company-name/type"/></xsl:attribute>
		<xsl:attribute name="VALUE"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/company-name/value"/></xsl:attribute>
	</input>
	
	<select>
		<xsl:attribute name="ID"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/groups/id"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/groups/name"/></xsl:attribute>
		<xsl:attribute name="MULTIPLE"><xsl:value-of select="ui/subscribers/mainframe/newsubscriber/groups/multiple"/></xsl:attribute>
		<xsl:attribute name="VALUE">nav button has loaded</xsl:attribute>
		<!--//options could be written - or post populated, part of the onload if element exists - populate content//-->
	</select>
	
</form>

</body>
</html>

</xsl:template>
</xsl:stylesheet>