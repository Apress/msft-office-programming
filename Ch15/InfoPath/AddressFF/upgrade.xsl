<?xml version="1.0" encoding="UTF-8" standalone="no"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:my="http://schemas.microsoft.com/office/infopath/2003/myXSD/2003-08-08T15:13:41" version="1.0">
	<xsl:output encoding="UTF-8" method="xml"/>
	<xsl:template match="/">
		<xsl:copy-of select="processing-instruction() | comment()"/>
		<xsl:choose>
			<xsl:when test="my:myFields">
				<xsl:apply-templates select="my:myFields" mode="_0"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:variable name="var">
					<xsl:element name="my:myFields"/>
				</xsl:variable>
				<xsl:apply-templates select="msxsl:node-set($var)/*" mode="_0"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template match="my:myFields" mode="_0">
		<xsl:copy>
			<xsl:element name="my:txtFirstName">
				<xsl:copy-of select="my:txtFirstName/text()[1]"/>
			</xsl:element>
			<xsl:element name="my:txtLastName">
				<xsl:copy-of select="my:txtLastName/text()[1]"/>
			</xsl:element>
			<xsl:element name="my:txtStreet">
				<xsl:copy-of select="my:txtStreet/text()[1]"/>
			</xsl:element>
			<xsl:element name="my:txtCity">
				<xsl:copy-of select="my:txtCity/text()[1]"/>
			</xsl:element>
			<xsl:element name="my:cboState">
				<xsl:choose>
					<xsl:when test="my:cboState">
						<xsl:copy-of select="my:cboState/text()[1]"/>
					</xsl:when>
					<xsl:otherwise>AZ</xsl:otherwise>
				</xsl:choose>
			</xsl:element>
			<xsl:element name="my:txtZip">
				<xsl:copy-of select="my:txtZip/text()[1]"/>
			</xsl:element>
		</xsl:copy>
	</xsl:template>
</xsl:stylesheet>