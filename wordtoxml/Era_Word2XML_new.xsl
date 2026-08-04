<?xml version="1.0"?>
<xsl:stylesheet
    xmlns:xsl='http://www.w3.org/1999/XSL/Transform'
    xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
    version='2.0' xmlns:o="http://schemas.microsoft.com/office/word/2003/wordml"
	xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint"
	exclude-result-prefixes="w wx vt my xs" xmlns:v="urn:schemas-microsoft-com:vml"
	xmlns:my="urn:my-custom-functions"
	xmlns:xs="http://www.w3.org/2001/XMLSchema" 
	xmlns:oasis="//OASIS//DTD XML Exchange Table Model 19990315//EN" xmlns:mml="http://www.w3.org/1998/Math/MathML" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:aml="http://schemas.microsoft.com/aml/2001/core">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="no"/>
	<xsl:include href="table.xsl"/>

	<xsl:variable name="allStyles" select="document('styles.xml', /)/w:styles"/>
	<xsl:variable name="docDefaultsRPr" select="$allStyles/w:docDefaults/w:rPrDefault/w:rPr"/>

	<!-- Walk basedOn chain root-first, return sequence of style elements -->
	<xsl:function name="my:style-chain" as="element(w:style)*">
		<xsl:param name="styleId" as="xs:string?"/>
		<xsl:if test="$styleId">
			<xsl:variable name="style" select="$allStyles/w:style[@w:styleId=$styleId]"/>
			<xsl:sequence select="my:style-chain($style/w:basedOn/@w:val)"/>
			<xsl:sequence select="$style"/>
		</xsl:if>
	</xsl:function>

	<!-- last occurrence of each rPr child element wins (proper override) -->
	<xsl:function name="my:merge-rpr" as="element()*">
		<xsl:param name="nodes" as="element()*"/>
		<xsl:for-each-group select="$nodes" group-by="name()">
			<xsl:sequence select="current-group()[last()]"/>
		</xsl:for-each-group>
	</xsl:function>

	<!-- fully resolved rPr for a given run: docDefaults -> pStyle chain -> rStyle chain -> direct rPr -->
	<xsl:function name="my:effective-rpr" as="element()*">
		<xsl:param name="run" as="element(w:r)"/>
		<xsl:variable name="pStyleId" select="$run/ancestor::w:p[1]/w:pPr/w:pStyle/@w:val"/>
		<xsl:variable name="rStyleId" select="$run/w:rPr/w:rStyle/@w:val"/>
		<xsl:sequence select="my:merge-rpr((
			$docDefaultsRPr/*,
			my:style-chain($pStyleId)/w:rPr/*,
			my:style-chain($rStyleId)/w:rPr/*,
			$run/w:rPr/*
		))"/>
	</xsl:function>

	<xsl:template match="w:document">
			<xsl:apply-templates select="./w:body"/>
	</xsl:template>
	
<xsl:template match="w:body">
	
	<!-- <xsl:text disable-output-escaping="yes">&lt;br/&gt;&lt;!DOCTYPE article PUBLIC "-//NLM//DTD Journal Archiving and Interchange DTD v2.2 20060430//EN" ""&gt;&lt;br/&gt;</xsl:text> -->

	<html><br/>
		<head><br/>
			<title></title><br/>
			<link rel="stylesheet" type="text/css" href="epub.css"/><br/>
		</head><br/>

		<body><br/>
			<xsl:apply-templates/>
		</body><br/>
	</html>
			
</xsl:template>

<xsl:template match="w:comments">
	<br/><comments><xsl:apply-templates/></comments>
</xsl:template>

<xsl:template match="w:object">
<object/>
</xsl:template>

<xsl:template match="w:drawing">
<object/>
</xsl:template>

<xsl:template match="v:shape">
<object/>
</xsl:template>
<!--
<xsl:template match="/">
<root><xsl:apply-templates/></root>
</xsl:template>
-->
<xsl:template match="w:del/w:delText">
<del><xsl:apply-templates/></del>
</xsl:template>

<xsl:template match="w:ins/w:r/w:t">
<ins><xsl:apply-templates/></ins>
</xsl:template>

<xsl:template match="w:p">
	<xsl:choose>
<!-- 	
		<xsl:when test="w:pPr/w:pStyle/@w:val">
			<xsl:element name="{w:pPr/w:pStyle/@w:val}"><xsl:apply-templates/></xsl:element><br/>
		</xsl:when>
 -->
		<xsl:when test="w:pPr/w:pStyle/@w:val">
			<p>
				<xsl:attribute name="class"><xsl:value-of select="w:pPr/w:pStyle/@w:val"/></xsl:attribute>
				<xsl:apply-templates/>
			</p><br/>
		</xsl:when>

		<xsl:when test="w:pPr/w:pStyle/@w:val='TableHeadA'">
			<xsl:apply-templates/>
		</xsl:when>
		<xsl:when test="w:pPr/w:pStyle/@w:val='BoxHeadA'">
			<boxtitle1><xsl:apply-templates/></boxtitle1>
		</xsl:when>
		<xsl:when test="contains(w:r[1]/w:instrText, 'List-Start')">
			<list-start/><br/>
		</xsl:when>
		<xsl:when test="contains(w:r[1]/w:instrText, 'List-End')">
			<list-end/><br/>
		</xsl:when>
		<xsl:when test="w:pPr/w:listPr">
			<xsl:choose>
				<xsl:when test="w:pPr/w:listPr/w:ilvl/@w:val='0'">
					<item1><label><xsl:value-of select="w:pPr/w:listPr/wx:t/@wx:val"/></label>  <xsl:apply-templates/></item1><br/>
				</xsl:when>
				<xsl:when test="w:pPr/w:listPr/w:ilvl/@w:val='1'">
					<item2><label><xsl:value-of select="w:pPr/w:listPr/wx:t/@wx:val"/></label>  <xsl:apply-templates/></item2><br/>
				</xsl:when>
				<xsl:when test="w:pPr/w:listPr/w:ilvl/@w:val='2'">
					<item3><label><xsl:value-of select="w:pPr/w:listPr/wx:t/@wx:val"/></label>  <xsl:apply-templates/></item3><br/>
				</xsl:when>
				<xsl:when test="w:pPr/w:listPr/w:ilvl/@w:val='3'">
					<item4><label><xsl:value-of select="w:pPr/w:listPr/wx:t/@wx:val"/></label>  <xsl:apply-templates/></item4><br/>
				</xsl:when>
				<xsl:when test="w:pPr/w:listPr/w:ilvl/@w:val='4'">
					<item5><label><xsl:value-of select="w:pPr/w:listPr/wx:t/@wx:val"/></label>  <xsl:apply-templates/></item5><br/>
				</xsl:when>
				<xsl:when test="w:pPr/w:listPr/w:ilvl/@w:val='5'">
					<item6><label><xsl:value-of select="w:pPr/w:listPr/wx:t/@wx:val"/></label>  <xsl:apply-templates/></item6><br/>
				</xsl:when>
				<xsl:otherwise>
					<item><xsl:apply-templates/></item><br/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:when>

		<xsl:otherwise>
			<p><xsl:apply-templates/></p><br/>
		</xsl:otherwise>
	</xsl:choose>


	
<!--Hidden Text-->
<!-- 	<xsl:choose>

		<xsl:when test="w:r/w:fldChar[@w:fldCharType='begin']">
			<xsl:text>&lt;hidden&gt;</xsl:text>
		</xsl:when>
		<xsl:when test="w:r/w:fldChar[@w:fldCharType='end']">
			<xsl:text>&lt;/hidden&gt;</xsl:text>
		</xsl:when>
	</xsl:choose> -->

</xsl:template>
<!--Comment starts ends-->
<xsl:template match="w:comment">
	<br/><xsl:element name="chapcomment"><xsl:attribute name="id"><xsl:value-of select="@w:id"/></xsl:attribute><br/><xsl:apply-templates/></xsl:element><br/>
</xsl:template>

<xsl:template match="w:commentRangeStart">
	<xsl:text disable-output-escaping="yes">&lt;comment id="</xsl:text>
	<xsl:value-of select="@w:id"/>
	<xsl:text disable-output-escaping="yes">"&gt;</xsl:text>
</xsl:template>
<xsl:template match="w:commentRangeEnd">
		<xsl:text disable-output-escaping="yes">&lt;/comment&gt;</xsl:text>
</xsl:template>


<xsl:template match="//w:tab">
	<tab/>
</xsl:template>

	
<xsl:template match="w:r">
		<xsl:variable name="eff" select="my:effective-rpr(.)"/>
		<xsl:choose>

			<xsl:when test="$eff[self::w:b][not(@w:val='off')] and $eff[self::w:i][not(@w:val='off')] and $eff[self::w:vertAlign]/@w:val='superscript' and $eff[self::w:u]/@w:val='single' and w:rPr/w:rStyle/@w:val">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><xsl:element name="{w:rPr/w:rStyle/@w:val}"><sup><strong><em><u><xsl:apply-templates/></u></em></strong></sup></xsl:element><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:b][not(@w:val='off')] and $eff[self::w:i][not(@w:val='off')] and $eff[self::w:vertAlign]/@w:val='subscript' and $eff[self::w:u]/@w:val='single' and w:rPr/w:rStyle/@w:val">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><xsl:element name="{w:rPr/w:rStyle/@w:val}"><sub><strong><em><u><xsl:apply-templates/></u></em></strong></sub></xsl:element><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:b][not(@w:val='off')] and $eff[self::w:i][not(@w:val='off')] and $eff[self::w:vertAlign]/@w:val='superscript' and $eff[self::w:u]/@w:val='single'">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><sup><strong><em><u><xsl:apply-templates/></u></em></strong></sup><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:b][not(@w:val='off')] and $eff[self::w:i][not(@w:val='off')] and $eff[self::w:vertAlign]/@w:val='subscript' and $eff[self::w:u]/@w:val='single'">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><sub><strong><em><u><xsl:apply-templates/></u></em></strong></sub><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:b][not(@w:val='off')] and $eff[self::w:i][not(@w:val='off')] and $eff[self::w:vertAlign]/@w:val='superscript' and w:rPr/w:rStyle/@w:val">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><xsl:element name="{w:rPr/w:rStyle/@w:val}"><sup><strong><em><xsl:apply-templates/></em></strong></sup></xsl:element><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:b][not(@w:val='off')] and $eff[self::w:i][not(@w:val='off')] and $eff[self::w:vertAlign]/@w:val='subscript' and w:rPr/w:rStyle/@w:val">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><xsl:element name="{w:rPr/w:rStyle/@w:val}"><sub><strong><em><xsl:apply-templates/></em></strong></sub></xsl:element><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:b][not(@w:val='off')] and $eff[self::w:i][not(@w:val='off')] and $eff[self::w:vertAlign]/@w:val='superscript'">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><sup><strong><em><xsl:apply-templates/></em></strong></sup><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:b][not(@w:val='off')] and $eff[self::w:i][not(@w:val='off')] and $eff[self::w:vertAlign]/@w:val='subscript'">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><sub><strong><em><xsl:apply-templates/></em></strong></sub><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:b][not(@w:val='off')] and $eff[self::w:i][not(@w:val='off')] and $eff[self::w:u]/@w:val='single' and w:rPr/w:rStyle/@w:val">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><xsl:element name="{w:rPr/w:rStyle/@w:val}"><strong><em><u><xsl:apply-templates/></u></em></strong></xsl:element><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:b][not(@w:val='off')] and $eff[self::w:i][not(@w:val='off')] and $eff[self::w:u]/@w:val='single'">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><strong><em><u><xsl:apply-templates/></u></em></strong><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:b][not(@w:val='off')] and $eff[self::w:i][not(@w:val='off')] and w:rPr/w:rStyle/@w:val">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><xsl:element name="{w:rPr/w:rStyle/@w:val}"><strong><em><xsl:apply-templates/></em></strong></xsl:element><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:b][not(@w:val='off')] and $eff[self::w:i][not(@w:val='off')]">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><strong><em><xsl:apply-templates/></em></strong><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:b][not(@w:val='off')] and $eff[self::w:u]/@w:val='single' and w:rPr/w:rStyle/@w:val">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><xsl:element name="{w:rPr/w:rStyle/@w:val}"><strong><u><xsl:apply-templates/></u></strong></xsl:element><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:b][not(@w:val='off')] and $eff[self::w:u]/@w:val='single'">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><strong><u><xsl:apply-templates/></u></strong><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:i][not(@w:val='off')] and $eff[self::w:u]/@w:val='single' and w:rPr/w:rStyle/@w:val">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><xsl:element name="{w:rPr/w:rStyle/@w:val}"><em><u><xsl:apply-templates/></u></em></xsl:element><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:i][not(@w:val='off')] and $eff[self::w:u]/@w:val='single'">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><em><u><xsl:apply-templates/></u></em><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:u]/@w:val='single' and $eff[self::w:vertAlign]/@w:val='superscript' and w:rPr/w:rStyle/@w:val">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><xsl:element name="{w:rPr/w:rStyle/@w:val}"><sup><u><xsl:apply-templates/></u></sup></xsl:element><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:u]/@w:val='single' and $eff[self::w:vertAlign]/@w:val='superscript'">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><sup><u><xsl:apply-templates/></u></sup><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:b][not(@w:val='off')] and $eff[self::w:vertAlign]/@w:val='superscript' and w:rPr/w:rStyle/@w:val">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><xsl:element name="{w:rPr/w:rStyle/@w:val}"><sup><strong><xsl:apply-templates/></strong></sup></xsl:element><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:b][not(@w:val='off')] and $eff[self::w:vertAlign]/@w:val='superscript'">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><sup><strong><xsl:apply-templates/></strong></sup><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:i][not(@w:val='off')] and $eff[self::w:vertAlign]/@w:val='superscript' and w:rPr/w:rStyle/@w:val">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><xsl:element name="{w:rPr/w:rStyle/@w:val}"><sup><em><xsl:apply-templates/></em></sup></xsl:element><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:i][not(@w:val='off')] and $eff[self::w:vertAlign]/@w:val='superscript'">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><sup><em><xsl:apply-templates/></em></sup><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:u]/@w:val='single' and $eff[self::w:vertAlign]/@w:val='subscript' and w:rPr/w:rStyle/@w:val">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><xsl:element name="{w:rPr/w:rStyle/@w:val}"><sub><u><xsl:apply-templates/></u></sub></xsl:element><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:u]/@w:val='single' and $eff[self::w:vertAlign]/@w:val='subscript'">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><sub><u><xsl:apply-templates/></u></sub><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:b][not(@w:val='off')] and $eff[self::w:vertAlign]/@w:val='subscript' and w:rPr/w:rStyle/@w:val">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><xsl:element name="{w:rPr/w:rStyle/@w:val}"><sub><strong><xsl:apply-templates/></strong></sub></xsl:element><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:b][not(@w:val='off')] and $eff[self::w:vertAlign]/@w:val='subscript'">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><sub><strong><xsl:apply-templates/></strong></sub><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:i][not(@w:val='off')] and $eff[self::w:vertAlign]/@w:val='subscript' and w:rPr/w:rStyle/@w:val">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><xsl:element name="{w:rPr/w:rStyle/@w:val}"><sub><em><xsl:apply-templates/></em></sub></xsl:element><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:i][not(@w:val='off')] and $eff[self::w:vertAlign]/@w:val='subscript'">
				<sub><em><xsl:apply-templates/></em></sub>
			</xsl:when>
			<xsl:when test="$eff[self::w:b][not(@w:val='off')] and w:rPr/w:rStyle/@w:val">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><xsl:element name="{w:rPr/w:rStyle/@w:val}"><strong><xsl:apply-templates/></strong></xsl:element><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:b][not(@w:val='off')]">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><strong><xsl:apply-templates/></strong><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:i][not(@w:val='off')] and w:rPr/w:rStyle/@w:val">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><xsl:element name="{w:rPr/w:rStyle/@w:val}"><em><xsl:apply-templates/></em></xsl:element><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:i][not(@w:val='off')]">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if>
				<xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if>
				<xsl:if test="w:instrText"><xsl:text disable-output-escaping="yes">&lt;hidden&gt;</xsl:text></xsl:if>
				<em><xsl:apply-templates/></em>
				<xsl:if test="w:instrText"><xsl:text disable-output-escaping="yes">&lt;/hidden&gt;</xsl:text></xsl:if>
				<xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if>
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:vertAlign]/@w:val='subscript' and w:rPr/w:rStyle/@w:val">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><xsl:element name="{w:rPr/w:rStyle/@w:val}"><sub><xsl:apply-templates/></sub></xsl:element><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:vertAlign]/@w:val='subscript'">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><sub><xsl:apply-templates/></sub><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:vertAlign]/@w:val='superscript' and w:rPr/w:rStyle/@w:val">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><xsl:element name="{w:rPr/w:rStyle/@w:val}"><sup><xsl:apply-templates/></sup></xsl:element><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:vertAlign]/@w:val='superscript'">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><sup><xsl:apply-templates/></sup><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:u]/@w:val='single' and w:rPr/w:rStyle/@w:val">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><xsl:element name="{w:rPr/w:rStyle/@w:val}"><u><xsl:apply-templates/></u></xsl:element><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="$eff[self::w:u]/@w:val='single'">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><u><xsl:apply-templates/></u><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:when test="w:rPr/w:rStyle/@w:val='CHead'">
				<title><xsl:apply-templates/></title>
			</xsl:when>
			<xsl:when test="w:rPr/w:rStyle/@w:val='ng-term' or w:rPr/w:rStyle/@w:val='term'">
				<xsl:element name="{w:rPr/w:rStyle/@w:val}"><xsl:apply-templates/></xsl:element> <!--bold removed on 25 Mar, 09 as per Vijay voice -->
			</xsl:when>
	<xsl:when test="w:rPr/w:rStyle/@w:val='Hyperlink' and $eff[self::w:vertAlign]/@w:val='superscript'">
				<sup><xsl:element name="{w:rPr/w:rStyle/@w:val}"><xsl:apply-templates/></xsl:element></sup>
			</xsl:when> 
			<xsl:when test="w:rPr/w:rStyle/@w:val">
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if><xsl:element name="{w:rPr/w:rStyle/@w:val}"><xsl:apply-templates/></xsl:element><xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if><xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:when>
			<xsl:otherwise>
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;smallcaps&gt;</xsl:text></xsl:if>
				<xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;strikethrough&gt;</xsl:text></xsl:if>
				<xsl:apply-templates/>
				<xsl:if test="$eff[self::w:strike][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/strikethrough&gt;</xsl:text></xsl:if>
				<xsl:if test="$eff[self::w:smallCaps][not(@w:val='off')]"><xsl:text disable-output-escaping="yes">&lt;/smallcaps&gt;</xsl:text></xsl:if>
			</xsl:otherwise>
		</xsl:choose>
<!-- 		<xsl:choose>
			<xsl:when test="w:instrText">
			<hidden><xsl:apply-templates/></hidden>
			</xsl:when>
		</xsl:choose>
 -->
	</xsl:template>

	
	<xsl:template match="w:hlink">
		<xsl:choose>
			<xsl:when test="@w:bookmark and w:r/w:rPr/w:vertAlign/@w:val='superscript'">
				<sup><xsl:element name="xref"><xsl:attribute name="href"><xsl:value-of select="@w:bookmark"/></xsl:attribute><xsl:apply-templates/></xsl:element></sup>
			</xsl:when>
			<xsl:when test="@w:bookmark">
				<xsl:element name="xref"><xsl:attribute name="href"><xsl:value-of select="@w:bookmark"/></xsl:attribute><xsl:apply-templates/></xsl:element>
			</xsl:when>
			<xsl:otherwise>
				<xref><xsl:apply-templates/></xref>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	
</xsl:stylesheet>

