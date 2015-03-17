<?xml version="1.0"?>
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:output method="text" />

<xsl:template match="/">

<!--// BOUNCE DIAGNOSTIC MESSAGE REMOVED FROM LISTING AS CURRENT VIEW DOES NOT CONTAIN IT UNLESS VIEWING BOUNCES //-->
<xsl:text>Email Address,Successfully Sent,Number of Times Opened,First Opened,Clicked,First Clicked,Bounced,Bounce Type,Bounce Message</xsl:text> 


<!--// IF ITS A URL SUMMARY//-->
<xsl:text>,URL</xsl:text> 
<xsl:text>,Number of clicks</xsl:text>

<!--//ADD IN THE EMAIL ADDRESS//-->
<xsl:text>,Title</xsl:text> 
<xsl:text>,First_Name</xsl:text> 
<xsl:text>,Last_Name</xsl:text> 
<xsl:text>,Email List</xsl:text> 
<xsl:text>,Email Type</xsl:text> 
<xsl:text>,Mobile Phone</xsl:text> 
<xsl:text>,Company Name</xsl:text> 
<xsl:text>,Date Added</xsl:text> 
<xsl:text>,Source</xsl:text> 
<xsl:text>,Use</xsl:text> 
<xsl:text>,Unsubscribed</xsl:text>
<xsl:text>,Date Unsubscribed</xsl:text>

<!--//ADD IN THE ADDRESS TITLE//-->
<xsl:text>,Address1</xsl:text> 
<xsl:text>,Address2</xsl:text>
<xsl:text>,Address3</xsl:text>
<xsl:text>,City</xsl:text>
<xsl:text>,County</xsl:text>
<xsl:text>,Country</xsl:text>
<xsl:text>,TelNo</xsl:text>
<xsl:text>,FaxNo</xsl:text>
<xsl:text>,Postcode</xsl:text>

<!--//ADD IN THE CUSTOM FIELDS TITLE//-->
<xsl:for-each select="Export/View_ReportSummary[1]/EmailAddresses/Addresses/View_CUSTOMFIELDS">
	<xsl:text>,</xsl:text>
	<xsl:value-of select="translate(string(Custom_Field_Name),',',';')" />
</xsl:for-each>

<xsl:text>&#xa;</xsl:text>

<!--// TITLES PRODUCED now PRODUCE DATA//-->

<xsl:for-each select="Export/View_ReportSummary">

	<xsl:variable name="totalCount" select="0"/>
	
	<xsl:variable name="od" select="OpenedDateTime"/>
	<xsl:variable name="odYear" select="substring($od,1,4)"/>
	<xsl:variable name="odMonth" select="substring($od,6,2)"/>
	<xsl:variable name="odDay" select="substring($od,9,2)"/>
	<xsl:variable name="odHour" select="substring($od,12,2)"/>
	<xsl:variable name="odMin" select="substring($od,15,2)"/>
	<xsl:variable name="odSec" select="substring($od,18,2)"/>

	<xsl:variable name="cd" select="ClickedDateTime"/>
	<xsl:variable name="cdYear" select="substring($cd,1,4)"/>
	<xsl:variable name="cdMonth" select="substring($cd,6,2)"/>
	<xsl:variable name="cdDay" select="substring($cd,9,2)"/>
	<xsl:variable name="cdHour" select="substring($cd,12,2)"/>
	<xsl:variable name="cdMin" select="substring($cd,15,2)"/>
	<xsl:variable name="cdSec" select="substring($cd,18,2)"/>

  <!--//DATA STARTS HERE//-->
  
  	<xsl:value-of select="translate(EmailAddress,',',';')" /><xsl:text>,</xsl:text>
	
	<xsl:if test="SuccessfullySent[.=1]">yes</xsl:if><xsl:if test="SuccessfullySent[.=0]">no</xsl:if><xsl:text>,</xsl:text>
	
	<xsl:value-of select="NoTimesOpened" /><xsl:text>,</xsl:text>
	
	<xsl:if test="($odYear = 1900)">,</xsl:if><xsl:if test="($odYear > 1901)"><xsl:value-of select="$odDay" />/<xsl:value-of select="$odMonth" />/<xsl:value-of select="$odYear" /> - <xsl:value-of select="$odHour" />:<xsl:value-of select="$odMin" />:<xsl:value-of select="$odSec" /><xsl:text>,</xsl:text>
	
	</xsl:if><xsl:if test="($cdYear = 1900)">no,</xsl:if>
	<xsl:if test="($cdYear > 1901)">yes,<xsl:value-of select="$cdDay" />/<xsl:value-of select="$cdMonth" />/<xsl:value-of select="$cdYear" /> - <xsl:value-of select="$cdHour" />:<xsl:value-of select="$cdMin" />:<xsl:value-of select="$cdSec" /></xsl:if>
	
	<xsl:text>,</xsl:text>
	<xsl:if test="Bounced[.=1]">yes</xsl:if><xsl:if test="Bounced[.=0]">no</xsl:if><xsl:text>,</xsl:text>
	
	<xsl:if test="Bounced[.=1]">
		<xsl:choose>
		  <xsl:when test="BouncedType[.=0]">
			<xsl:text>SOFT: Generic Bounce</xsl:text>
		  </xsl:when>
		  <xsl:when test="BouncedType[.=1]">
			<xsl:text>HARD: AOL User Unknown</xsl:text>
		  </xsl:when>
		  <xsl:when test="BouncedType[.=2]">
			<xsl:text>HARD: Hotmail User Unknown</xsl:text>
		  </xsl:when>
		  <xsl:when test="BouncedType[.=3]">
			<xsl:text>HARD: User Unknown</xsl:text>
		  </xsl:when>
		  <xsl:when test="BouncedType[.=4]">
			<xsl:text>HARD: Recipient Unknown</xsl:text>
		  </xsl:when>
		  <xsl:when test="BouncedType[.=5]">
			<xsl:text>HARD: Yahoo User Unknown</xsl:text>
		  </xsl:when>
		  <xsl:when test="BouncedType[.=10]">
			<xsl:text>HARD: Host Unknown</xsl:text>
		  </xsl:when>
		  <xsl:when test="BouncedType[.=11]">
			<xsl:text>HARD: No Such Domain</xsl:text>
		  </xsl:when>
		  
		  <xsl:otherwise>
			<xsl:text>GENERIC BOUNCE</xsl:text>
		  </xsl:otherwise>
		</xsl:choose>
		<xsl:text>,</xsl:text>
	</xsl:if>
	<xsl:if test="Bounced[.=0]">
		<xsl:text>,</xsl:text>
	</xsl:if>
	<xsl:if test="Bounced[.=1]">
		<xsl:if test="string-length(strBounceMessage) = 0">
			<xsl:text>For full bounce message export bounces</xsl:text>
		</xsl:if>
	</xsl:if>
	<xsl:value-of select="translate(translate(strBounceMessage,'&#x0A;',' '),',',';')" /><xsl:text>,</xsl:text>
	
	<!--// URL CLICKS //-->
	<xsl:value-of select="translate(Url,',',';')" /><xsl:text>,</xsl:text>
	<xsl:value-of select="translate(NoTimesClicked,',',';')" /><xsl:text>,</xsl:text>
	
	<!--//Emailaddress DATA//-->
	<xsl:for-each select="EmailAddresses">
	
		<xsl:variable name="ad" select="DateAdded"/>
		<xsl:variable name="adYear" select="substring($ad,1,4)"/>
		<xsl:variable name="adMonth" select="substring($ad,6,2)"/>
		<xsl:variable name="adDay" select="substring($ad,9,2)"/>
		<xsl:variable name="adHour" select="substring($ad,12,2)"/>
		<xsl:variable name="adMin" select="substring($ad,15,2)"/>
		<xsl:variable name="adSec" select="substring($ad,18,2)"/>
		
		<xsl:variable name="ud" select="UnsubscribeDate"/>
		<xsl:variable name="udYear" select="substring($ud,1,4)"/>
		<xsl:variable name="udMonth" select="substring($ud,6,2)"/>
		<xsl:variable name="udDay" select="substring($ud,9,2)"/>
		<xsl:variable name="udHour" select="substring($ud,12,2)"/>
		<xsl:variable name="udMin" select="substring($ud,15,2)"/>
		<xsl:variable name="udSec" select="substring($ud,18,2)"/>
	
		<xsl:value-of select="translate(Subscriber_Title,',',';')" /><xsl:text>,</xsl:text>
		<xsl:value-of select="translate(Subscriber_First_Name,',',';')" /><xsl:text>,</xsl:text>
		<xsl:value-of select="translate(Subscriber_Last_Name,',',';')" /><xsl:text>,</xsl:text>
		<xsl:if test="EmailList[.=1]">yes</xsl:if><xsl:if test="EmailList[.=0]">no</xsl:if><xsl:text>,</xsl:text>
		<xsl:if test="EmailType[.=2]">text</xsl:if><xsl:if test="EmailType[.=1]">html</xsl:if><xsl:if test="EmailType[.=0]">multi-part</xsl:if><xsl:text>,</xsl:text>
		<xsl:value-of select="translate(string(MobilePhone),',',';')" /><xsl:text>,</xsl:text>
		<xsl:value-of select="translate(CompanyName,',',';')" /><xsl:text>,</xsl:text>
	<xsl:if test="($adYear > 1901)"><xsl:value-of select="$adDay" />/<xsl:value-of select="$adMonth" />/<xsl:value-of select="$adYear" /> - <xsl:value-of select="$adHour" />:<xsl:value-of select="$adMin" />:<xsl:value-of select="$adSec" /></xsl:if>
		<xsl:text>,</xsl:text>
		<xsl:value-of select="translate(Source,',',';')" /><xsl:text>,</xsl:text>
		<xsl:if test="Use[.=1]">yes</xsl:if><xsl:if test="Use[.=0]">no</xsl:if><xsl:text>,</xsl:text>
		<xsl:if test="Unsubscribed[.=1]">yes</xsl:if><xsl:if test="Unsubscribed[.=0]">no</xsl:if><xsl:text>,</xsl:text>
		<xsl:if test="($udYear > 1901)"><xsl:value-of select="$udDay" />/<xsl:value-of select="$udMonth" />/<xsl:value-of select="$udYear" /> - <xsl:value-of select="$udHour" />:<xsl:value-of select="$udMin" />:<xsl:value-of select="$udSec" /></xsl:if>
		<xsl:text>,</xsl:text>
		
		
		
	</xsl:for-each>
	
	<!--//ADDRESS DATA//-->
	<xsl:for-each select="EmailAddresses/Addresses">
		
		<xsl:value-of select="translate(Address1,',',';')" /><xsl:text>,</xsl:text>
		<xsl:value-of select="translate(Address2,',',';')" /><xsl:text>,</xsl:text>
		<xsl:value-of select="translate(Address3,',',';')" /><xsl:text>,</xsl:text>
		<xsl:value-of select="translate(City,',',';')" /><xsl:text>,</xsl:text>
		<xsl:value-of select="translate(County,',',';')" /><xsl:text>,</xsl:text>
		<xsl:value-of select="translate(Country,',',';')" /><xsl:text>,</xsl:text>
		<xsl:value-of select="translate(string(TelNo),',',';')" /><xsl:text>,</xsl:text>
		<xsl:value-of select="translate(string(FaxNo),',',';')" /><xsl:text>,</xsl:text>
		<xsl:value-of select="translate(Postcode,',',';')" />
		
	</xsl:for-each>
	<!--//CUSTOMFIELDS DATA//-->
	<!--//WILL NEED TO OUTPUT DYNAMIC DATA FOR CUSTOM FIELDS - AND DOP A LOOPING COUNT!//-->
	
	<xsl:for-each select="EmailAddresses/Addresses/View_CUSTOMFIELDS">
		<xsl:text>,</xsl:text><xsl:value-of select="translate(string(Custom_Field_Data),',',';')" />
	</xsl:for-each>

	<xsl:text>&#xa;</xsl:text>

</xsl:for-each>





</xsl:template>
</xsl:stylesheet>