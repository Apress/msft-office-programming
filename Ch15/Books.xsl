<?xml version="1.0" ?>
<xsl:stylesheet version="1.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:output method="text" />

    <!-- Root -->
    <xsl:template match="/">
        <xsl:apply-templates />
    </xsl:template>

    <!-- Books -->
    <xsl:template match="Books">
        <xsl:apply-templates />
    </xsl:template>

    <!-- Book -->
    <xsl:template match="Book">
        <xsl:text>Title: </xsl:text>
        <xsl:value-of select="Title" />
        <xsl:text>
</xsl:text>

        <xsl:text>Price: </xsl:text>
        <xsl:value-of select="@Price" />
        <xsl:text>
</xsl:text>

        <xsl:text>Pages: </xsl:text>
        <xsl:value-of select="@Pages" />
        <xsl:text>

</xsl:text>
    </xsl:template>

</xsl:stylesheet>
