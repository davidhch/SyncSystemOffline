<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:output method="html" encoding="iso-8859-1" doctype-public="-//W3C//DTD XHTML 1.0 Transitional//EN" doctype-system="http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"/>
<xsl:template match="/">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>
<title>new subscriber</title>
<!--//note JS may be moved elsewhere, but for now it is here//-->
<script language="JavaScript" type="text/javascript" src="JSGoesHere.js"></script>
</head>
<body>

<form name="FormName" method="post" action="example.asp">

	<input>
		<xsl:attribute name="ID"><xsl:value-of select="ui/id102/id"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="ui/id102/name"/></xsl:attribute>
		<xsl:attribute name="TYPE"><xsl:value-of select="ui/id102/type"/></xsl:attribute>
	</input>
	
	<input>
		<xsl:attribute name="ID"><xsl:value-of select="ui/id103/id"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="ui/id103/name"/></xsl:attribute>
		<xsl:attribute name="TYPE"><xsl:value-of select="ui/id103/type"/></xsl:attribute>
		<xsl:attribute name="VALUE"><xsl:value-of select="ui/id103/value"/></xsl:attribute>
	</input>
	
	<input>
		<xsl:attribute name="ID"><xsl:value-of select="ui/id104/id"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="ui/id104/name"/></xsl:attribute>
		<xsl:attribute name="TYPE"><xsl:value-of select="ui/id104/type"/></xsl:attribute>
		<xsl:attribute name="VALUE"><xsl:value-of select="ui/id104/value"/></xsl:attribute>
	</input>
	
	<input>
		<xsl:attribute name="ID"><xsl:value-of select="ui/id105/id"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="ui/id105/name"/></xsl:attribute>
		<xsl:attribute name="TYPE"><xsl:value-of select="ui/id105/type"/></xsl:attribute>
		<xsl:attribute name="VALUE"><xsl:value-of select="ui/id105/value"/></xsl:attribute>
	</input>
	
	<input>
		<xsl:attribute name="ID"><xsl:value-of select="ui/id106/id"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="ui//id106/name"/></xsl:attribute>
		<xsl:attribute name="TYPE"><xsl:value-of select="ui//id106/type"/></xsl:attribute>
		<xsl:attribute name="VALUE"><xsl:value-of select="ui//id106/value"/></xsl:attribute>
	</input>
	
	<input>
		<xsl:attribute name="ID"><xsl:value-of select="ui/id107/id"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="ui/id107/name"/></xsl:attribute>
		<xsl:attribute name="TYPE"><xsl:value-of select="ui/id107/type"/></xsl:attribute>
		<xsl:attribute name="VALUE"><xsl:value-of select="ui/id107/value"/></xsl:attribute>
	</input>	
	
	<input>
		<xsl:attribute name="ID"><xsl:value-of select="ui/id108/id"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="ui/id108/name"/></xsl:attribute>
		<xsl:attribute name="TYPE"><xsl:value-of select="ui/id108/type"/></xsl:attribute>
		<xsl:attribute name="VALUE"><xsl:value-of select="ui/id108/value"/></xsl:attribute>
	</input>	
	
	<input>
		<xsl:attribute name="ID"><xsl:value-of select="ui/id109/id"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="ui/id109/name"/></xsl:attribute>
		<xsl:attribute name="TYPE"><xsl:value-of select="ui/id109/type"/></xsl:attribute>
		<xsl:attribute name="VALUE"><xsl:value-of select="ui/id109/value"/></xsl:attribute>
	</input>
	
	<select>
		<xsl:attribute name="ID"><xsl:value-of select="ui/id110/id"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="ui/id110/name"/></xsl:attribute>
		<xsl:attribute name="MULTIPLE"><xsl:value-of select="ui/id110/multiple"/></xsl:attribute>
		<!--//options could be written - or post populated, part of the onload if element exists - populate content//-->
	</select>
	
</form>

</body>
</html>

</xsl:template>
</xsl:stylesheet>