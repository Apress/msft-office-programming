<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:my="http://schemas.microsoft.com/office/infopath/2003/myXSD/2003-08-08T15:13:41" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:xd="http://schemas.microsoft.com/office/infopath/2003" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:xdExtension="http://schemas.microsoft.com/office/infopath/2003/xslt/extension" xmlns:xdXDocument="http://schemas.microsoft.com/office/infopath/2003/xslt/xDocument" xmlns:xdSolution="http://schemas.microsoft.com/office/infopath/2003/xslt/solution" xmlns:xdFormatting="http://schemas.microsoft.com/office/infopath/2003/xslt/formatting" xmlns:xdImage="http://schemas.microsoft.com/office/infopath/2003/xslt/xImage">
	<xsl:output method="html" indent="no"/>
	<xsl:template match="my:myFields">
		<html>
			<head>
				<style tableEditor="TableStyleRulesID">TABLE.xdLayout TD {
	BORDER-RIGHT: medium none; BORDER-TOP: medium none; BORDER-LEFT: medium none; BORDER-BOTTOM: medium none
}
TABLE {
	BEHAVIOR: url (#default#urn::tables/NDTable)
}
TABLE.msoUcTable TD {
	BORDER-RIGHT: 1pt solid; BORDER-TOP: 1pt solid; BORDER-LEFT: 1pt solid; BORDER-BOTTOM: 1pt solid
}
</style>
				<meta http-equiv="Content-Type" content="text/html"></meta>
				<style controlStyle="controlStyle">BODY{margin-left:21px;color:windowtext;background-color:window;layout-grid:none;} 		.xdListItem {display:inline-block;width:100%;vertical-align:text-top;} 		.xdListBox,.xdComboBox{margin:1px;} 		.xdInlinePicture{margin:1px; BEHAVIOR: url(#default#urn::xdPicture) } 		.xdLinkedPicture{margin:1px; BEHAVIOR: url(#default#urn::xdPicture) url(#default#urn::controls/Binder) } 		.xdSection{border:1pt solid #FFFFFF;margin:6px 0px 6px 0px;padding:1px 1px 1px 5px;} 		.xdRepeatingSection{border:1pt solid #FFFFFF;margin:6px 0px 6px 0px;padding:1px 1px 1px 5px;} 		.xdBehavior_Formatting {BEHAVIOR: url(#default#urn::controls/Binder) url(#default#Formatting);} 	 .xdBehavior_FormattingNoBUI{BEHAVIOR: url(#default#CalPopup) url(#default#urn::controls/Binder) url(#default#Formatting);} 	.xdExpressionBox{margin: 1px;padding:1px;word-wrap: break-word;text-overflow: ellipsis;overflow-x:hidden;}.xdBehavior_GhostedText,.xdBehavior_GhostedTextNoBUI{BEHAVIOR: url(#default#urn::controls/Binder) url(#default#TextField) url(#default#GhostedText);}	.xdBehavior_GTFormatting{BEHAVIOR: url(#default#urn::controls/Binder) url(#default#Formatting) url(#default#GhostedText);}	.xdBehavior_GTFormattingNoBUI{BEHAVIOR: url(#default#CalPopup) url(#default#urn::controls/Binder) url(#default#Formatting) url(#default#GhostedText);}	.xdBehavior_Boolean{BEHAVIOR: url(#default#urn::controls/Binder) url(#default#BooleanHelper);}	.xdBehavior_Select{BEHAVIOR: url(#default#urn::controls/Binder) url(#default#SelectHelper);}	.xdRepeatingTable{BORDER-TOP-STYLE: none; BORDER-RIGHT-STYLE: none; BORDER-LEFT-STYLE: none; BORDER-BOTTOM-STYLE: none; BORDER-COLLAPSE: collapse; WORD-WRAP: break-word;}.xdTextBox{display:inline-block;white-space:nowrap;text-overflow:ellipsis;;padding:1px;margin:1px;border: 1pt solid #dcdcdc;color:windowtext;background-color:window;overflow:hidden;text-align:left;} 		.xdRichTextBox{display:inline-block;;padding:1px;margin:1px;border: 1pt solid #dcdcdc;color:windowtext;background-color:window;overflow-x:hidden;word-wrap:break-word;text-overflow:ellipsis;text-align:left;font-weight:normal;font-style:normal;text-decoration:none;vertical-align:baseline;} 		.xdDTPicker{;display:inline;margin:1px;margin-bottom: 2px;border: 1pt solid #dcdcdc;color:windowtext;background-color:window;overflow:hidden;} 		.xdDTText{height:100%;width:100%;margin-right:22px;overflow:hidden;padding:0px;white-space:nowrap;} 		.xdDTButton{margin-left:-21px;height:18px;width:20px;behavior: url(#default#DTPicker);} 		.xdRepeatingTable TD {VERTICAL-ALIGN: top;}</style>
				<style languageStyle="languageStyle">BODY {
	FONT-SIZE: 10pt; FONT-FAMILY: Verdana
}
TABLE {
	FONT-SIZE: 10pt; FONT-FAMILY: Verdana
}
SELECT {
	FONT-SIZE: 10pt; FONT-FAMILY: Verdana
}
.optionalPlaceholder {
	PADDING-LEFT: 20px; FONT-WEIGHT: normal; FONT-SIZE: xx-small; BEHAVIOR: url(#default#xOptional); COLOR: #333333; FONT-STYLE: normal; FONT-FAMILY: Verdana; TEXT-DECORATION: none
}
.langFont {
	FONT-FAMILY: Verdana
}
</style>
			</head>
			<body>
				<div>
					<font size="1">
						<u>F</u>irst Name</font>
				</div>
				<div>
					<font size="1"><span class="xdTextBox" hideFocus="1" title="First name" accessKey="F" xd:binding="my:txtFirstName" tabIndex="0" xd:xctname="PlainText" xd:CtrlId="CTRL1" style="FONT-SIZE: x-small; WIDTH: 144px; HEIGHT: 20px">
							<xsl:value-of select="my:txtFirstName"/>
						</span>
					</font>
				</div>
				<div>
					<font size="1">
						<u>L</u>ast Name</font>
				</div>
				<div>
					<font size="1"><span class="xdTextBox" hideFocus="1" title="Last name" accessKey="L" xd:binding="my:txtLastName" tabIndex="0" xd:xctname="PlainText" xd:CtrlId="CTRL2" style="FONT-SIZE: x-small; WIDTH: 144px; HEIGHT: 20px">
							<xsl:value-of select="my:txtLastName"/>
						</span>
					</font>
				</div>
				<div>
					<font size="1">
						<u>S</u>treet</font>
				</div>
				<div><span class="xdTextBox" hideFocus="1" title="Street" accessKey="S" xd:binding="my:txtStree" tabIndex="0" xd:xctname="PlainText" xd:CtrlId="CTRL3" style="FONT-SIZE: x-small; WIDTH: 144px; HEIGHT: 20px">
						<xsl:value-of select="my:txtStree"/>
					</span>
					<br/>
					<font size="1">
						<u>C</u>ity</font>
				</div>
				<div>
					<font size="1"><span class="xdTextBox" hideFocus="1" title="City" accessKey="C" xd:binding="my:txtCity" tabIndex="0" xd:xctname="PlainText" xd:CtrlId="CTRL4" style="FONT-SIZE: x-small; WIDTH: 144px; HEIGHT: 20px">
							<xsl:value-of select="my:txtCity"/>
						</span>
					</font>
				</div>
				<div>
					<font size="1">S<u>t</u>ate            <u>Z</u>ip</font>
				</div>
				<div>
					<font size="1"><select class="xdComboBox xdBehavior_Select" title="State" accessKey="T" size="1" xd:binding="my:cboState" tabIndex="0" xd:xctname="DropDown" xd:CtrlId="CTRL5" xd:boundProp="value" style="FONT-SIZE: x-small; WIDTH: 62px">
							<xsl:attribute name="value">
								<xsl:value-of select="my:cboState"/>
							</xsl:attribute>
							<option value="AZ">
								<xsl:if test="my:cboState=&quot;AZ&quot;">
									<xsl:attribute name="selected">selected</xsl:attribute>
								</xsl:if>AZ</option>
							<option value="CA">
								<xsl:if test="my:cboState=&quot;CA&quot;">
									<xsl:attribute name="selected">selected</xsl:attribute>
								</xsl:if>CA</option>
							<option value="CO">
								<xsl:if test="my:cboState=&quot;CO&quot;">
									<xsl:attribute name="selected">selected</xsl:attribute>
								</xsl:if>CO</option>
							<option value="OR">
								<xsl:if test="my:cboState=&quot;OR&quot;">
									<xsl:attribute name="selected">selected</xsl:attribute>
								</xsl:if>OR</option>
							<option value="NM">
								<xsl:if test="my:cboState=&quot;NM&quot;">
									<xsl:attribute name="selected">selected</xsl:attribute>
								</xsl:if>NM</option>
							<option value="NV">
								<xsl:if test="my:cboState=&quot;NV&quot;">
									<xsl:attribute name="selected">selected</xsl:attribute>
								</xsl:if>NV</option>
							<option value="UT">
								<xsl:if test="my:cboState=&quot;UT&quot;">
									<xsl:attribute name="selected">selected</xsl:attribute>
								</xsl:if>UT</option>
							<option value="WA">
								<xsl:if test="my:cboState=&quot;WA&quot;">
									<xsl:attribute name="selected">selected</xsl:attribute>
								</xsl:if>WA</option>
						</select>   </font><span class="xdTextBox" hideFocus="1" title="Zip code" accessKey="Z" xd:binding="my:txtZip" tabIndex="0" xd:xctname="PlainText" xd:CtrlId="CTRL6" style="FONT-SIZE: x-small; WIDTH: 68px; HEIGHT: 20px">
						<xsl:value-of select="my:txtZip"/>
					</span>
				</div>
			</body>
		</html>
	</xsl:template>
</xsl:stylesheet>
