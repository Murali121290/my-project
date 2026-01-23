<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
xmlns:o="http://schemas.microsoft.com/office/word/2003/wordml"
xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint"
exclude-result-prefixes="w wx">

<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="no"/>

<xsl:template match="w:vMerge">
	<xsl:if test="@w:val='restart'">
		<xsl:variable name="column-number" select="count(.|ancestor::w:tc[1]/preceding-sibling::w:tc[not(w:tcPr/w:gridSpan)]) + sum(ancestor::w:tc[1]/preceding-sibling::w:tc/w:tcPr/w:gridSpan/@w:val)"/>
		<xsl:call-template name="cell.rowspaned">
			<xsl:with-param name="rows-spanned" select="0"/>
			<xsl:with-param name="current-cell" select="ancestor::w:tr[1]/following-sibling::w:tr[1]/w:tc[count(.|preceding-sibling::w:tc[not(w:tcPr/w:gridSpan)]) + sum(preceding-sibling::w:tc/w:tcPr/w:gridSpan/@w:val) = $column-number]"/>
			<xsl:with-param name="current-column" select="$column-number"/>
		</xsl:call-template>
	</xsl:if>
</xsl:template>

<xsl:template name="cell.rowspaned">
	<xsl:param name="rows-spanned"/>
	<xsl:param name="current-cell"/>
	<xsl:param name="current-column"/>
	<xsl:choose>
		<xsl:when test="$current-cell/w:tcPr/w:vMerge[not(@w:val='restart')]">
			<xsl:call-template name="cell.rowspaned">
				<xsl:with-param name="rows-spanned" select="$rows-spanned + 1"/>
				<xsl:with-param name="current-cell" select="$current-cell/ancestor::w:tr[1]/following-sibling::w:tr[1]/w:tc[count(.|preceding-sibling::w:tc[not(w:tcPr/w:gridSpan)]) + sum(preceding-sibling::w:tc/w:tcPr/w:gridSpan/@w:val) = $current-column]"/>
				<xsl:with-param name="current-column" select="$current-column"/>
			</xsl:call-template>
		</xsl:when>
		<xsl:otherwise>
			<xsl:attribute name="morerows">
				<xsl:value-of select="$rows-spanned"/>
			</xsl:attribute>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>


<xsl:template match="w:tbl">
<xsl:choose>
	<xsl:when test="w:tblPr/w:tblBorders/w:top/@w:val='none' and w:tblPr/w:tblBorders/w:left/@w:val='none' and w:tblPr/w:tblBorders/w:bottom/@w:val='none' and w:tblPr/w:tblBorders/w:right/@w:val='none'">
		<table frame="none" orient="port"/><br/>
	</xsl:when>
	<xsl:when test="w:tblPr/w:tblBorders/w:top/@w:val='none' and w:tblPr/w:tblBorders/w:left/@w:val='none' and w:tblPr/w:tblBorders/w:right/@w:val='none'">
		<table frame="bottom" orient="port"/><br/>
	</xsl:when>
	<xsl:when test="w:tblPr/w:tblBorders/w:left/@w:val='none' and w:tblPr/w:tblBorders/w:bottom/@w:val='none' and w:tblPr/w:tblBorders/w:right/@w:val='none'">
		<table frame="top" orient="port"/><br/>
	</xsl:when>
	<xsl:when test="w:tblPr/w:tblBorders/w:left/@w:val='none' and w:tblPr/w:tblBorders/w:right/@w:val='none'">
		<table frame="topbot" orient="port"/><br/>
	</xsl:when>
	<xsl:otherwise>
		<table frame="all" orient="port"/><br/>
	</xsl:otherwise>
</xsl:choose>
<!-- Tgroup & ColSpec -->
<xsl:choose>
	<xsl:when test="w:tblPr/w:tblBorders/w:insideH/@w:val='none' and w:tblPr/w:tblBorders/w:insideV/@w:val='none'">
		<tgroup cols="{count(w:tblGrid/w:gridCol)}"/><br/>
	</xsl:when>
	<xsl:when test="w:tblPr/w:tblBorders/w:insideH/@w:val='none'">
		<tgroup colsep="1" rowsep="0" cols="{count(w:tblGrid/w:gridCol)}"/><br/>
	</xsl:when>
	<xsl:when test="w:tblPr/w:tblBorders/w:insideV/@w:val='none'">
		<tgroup colsep="0" rowsep="1" cols="{count(w:tblGrid/w:gridCol)}"/><br/>
	</xsl:when>
	<xsl:otherwise>
		<tgroup colsep="0" rowsep="1" cols="{count(w:tblGrid/w:gridCol)}"/><br/>
	</xsl:otherwise>
</xsl:choose>
<xsl:for-each select="w:tblGrid/w:gridCol">
	<colspec align="left" colname="col{position()}" colnum="{position()}" colwidth="{format-number(@w:w div 57, '##')}*"/><br/>
</xsl:for-each>
<!-- End -->

<xsl:for-each select="w:tr">
	<xsl:choose>
		<xsl:when test="count(*/*/*/w:t)='0'">
			<xsl:apply-templates/>
		</xsl:when>
		<xsl:when test="w:trPr/w:tblHeader">
			<thead><br/>
				<row><br/>
					<xsl:call-template name="myCells">
						<xsl:with-param name="n" select="1"/>
					</xsl:call-template>
				</row><br/>
			</thead><br/>
		</xsl:when>
		<xsl:otherwise>
			<tbody><br/>
			<row><br/>
				<xsl:call-template name="myCells">
					<xsl:with-param name="n" select="1"/>
				</xsl:call-template>
			</row><br/></tbody><br/>
		</xsl:otherwise>
	</xsl:choose>
</xsl:for-each>
<tgroup/><br/>
<table/><br/>
</xsl:template>

<xsl:template name="myCells">
<xsl:param name="n"/>
<xsl:param name="noCol" select="count(w:tc)+1"/>
<xsl:param name="Col" select="1"/>
	<xsl:if test="$noCol &gt; $Col">
	<xsl:for-each select="w:tc[$Col]">
		<xsl:choose>
			<xsl:when test="w:tcPr/w:vMerge/@w:val='restart'">
				<xsl:element name="entry">
				<xsl:choose>
						<xsl:when test="w:tcPr/w:gridSpan/@w:val">
							<xsl:attribute name="colname">col<xsl:value-of select="$n"/></xsl:attribute>
							<xsl:attribute name="namest">col<xsl:value-of select="$n"/></xsl:attribute>
							<xsl:attribute name="nameend">col<xsl:value-of select="$n + (w:tcPr/w:gridSpan/@w:val - 1)"/></xsl:attribute>
						</xsl:when>
						<xsl:otherwise>
							<xsl:attribute name="colname">col<xsl:value-of select="$n"/></xsl:attribute>
						</xsl:otherwise>
					</xsl:choose>
					<xsl:attribute name="morerows"><xsl:value-of select="@rowspan"/></xsl:attribute>
				<xsl:choose>
					<xsl:when test="w:tcPr/w:vAlign/@w:val='center'">
						<xsl:attribute name="valign">middle</xsl:attribute>
					</xsl:when>
					<xsl:when test="w:tcPr/w:vAlign/@w:val='bottom'">
						<xsl:attribute name="valign">bottom</xsl:attribute>
					</xsl:when>
					<xsl:otherwise>
						<xsl:attribute name="valign">top</xsl:attribute>
					</xsl:otherwise>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="w:tcPr/w:shd/@w:fill='C0C0C0'">
						<xsl:attribute name="align">char</xsl:attribute>
						<xsl:attribute name="char">dot</xsl:attribute>
					</xsl:when>
					<xsl:when test="w:p/w:pPr/w:jc">
						<xsl:attribute name="align"><xsl:value-of select="w:p/w:pPr/w:jc/@w:val"/></xsl:attribute>
					</xsl:when>
					<xsl:otherwise>
						<xsl:attribute name="align">left</xsl:attribute>					
					</xsl:otherwise>
				</xsl:choose>
				<xsl:apply-templates/>
			</xsl:element><br/>
			</xsl:when>
			<xsl:when test="w:tcPr/w:vMerge">
				<!--<morerows value="col{$n}"/><br/>
				<xsl:apply-templates/>-->
			</xsl:when>
			<xsl:when test="w:tcPr/w:gridSpan/@w:val">
				<entry namest="col{$n}" nameend="col{$n + w:tcPr/w:gridSpan/@w:val - 1}">
				<xsl:choose>
					<xsl:when test="w:tcPr/w:vAlign/@w:val='center'">
						<xsl:attribute name="valign">middle</xsl:attribute>
					</xsl:when>
					<xsl:when test="w:tcPr/w:vAlign/@w:val='bottom'">
						<xsl:attribute name="valign">bottom</xsl:attribute>
					</xsl:when>
					<xsl:otherwise>
						<xsl:attribute name="valign">top</xsl:attribute>			
					</xsl:otherwise>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="w:tcPr/w:shd/@w:fill='C0C0C0'">
						<xsl:attribute name="align">char</xsl:attribute>
						<xsl:attribute name="char">dot</xsl:attribute>
					</xsl:when>
					<xsl:when test="w:p/w:pPr/w:jc">
						<xsl:attribute name="align"><xsl:value-of select="w:p/w:pPr/w:jc/@w:val"/></xsl:attribute>
					</xsl:when>
					<xsl:otherwise>
						<xsl:attribute name="align">left</xsl:attribute>			
					</xsl:otherwise>
				</xsl:choose>
					<xsl:apply-templates/>
				</entry><br/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:element name="entry">
				<xsl:attribute name="colname">col<xsl:value-of select="$n"/></xsl:attribute>
				<xsl:choose>
					<xsl:when test="w:tcPr/w:tcBorders/w:bottom/@w:val='nil' and w:tcPr/w:tcBorders/w:right/@w:val='nil'">
						<xsl:attribute name="rowsep">0</xsl:attribute>
						<xsl:attribute name="colsep">0</xsl:attribute>
					</xsl:when>
					<xsl:when test="w:tcPr/w:tcBorders/w:bottom/@w:val='nil' and w:tcPr/w:tcBorders/w:right/@w:val">
						<xsl:attribute name="rowsep">0</xsl:attribute>
						<xsl:attribute name="colsep">1</xsl:attribute>
					</xsl:when>
					<xsl:when test="w:tcPr/w:tcBorders/w:bottom/@w:val and w:tcPr/w:tcBorders/w:right/@w:val='nil'">
						<xsl:attribute name="rowsep">1</xsl:attribute>
						<xsl:attribute name="colsep">0</xsl:attribute>
					</xsl:when>
					<xsl:when test="w:tcPr/w:tcBorders/w:bottom/@w:val and w:tcPr/w:tcBorders/w:right/@w:val">
						<xsl:attribute name="rowsep">1</xsl:attribute>
						<xsl:attribute name="colsep">1</xsl:attribute>
					</xsl:when>
					<xsl:when test="w:tcPr/w:tcBorders/w:bottom/@w:val='nil'">
						<xsl:attribute name="rowsep">0</xsl:attribute>
					</xsl:when>
					<xsl:when test="w:tcPr/w:tcBorders/w:right/@w:val='nil'">
						<xsl:attribute name="colsep">0</xsl:attribute>
					</xsl:when>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="w:tcPr/w:vAlign/@w:val='center'">
						<xsl:attribute name="valign">middle</xsl:attribute>
					</xsl:when>
					<xsl:when test="w:tcPr/w:vAlign/@w:val='bottom'">
						<xsl:attribute name="valign">bottom</xsl:attribute>
					</xsl:when>
					<xsl:otherwise>
						<xsl:attribute name="valign">top</xsl:attribute>			
					</xsl:otherwise>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="w:tcPr/w:shd/@w:fill='C0C0C0'">
						<xsl:attribute name="align">char</xsl:attribute>
						<xsl:attribute name="char">dot</xsl:attribute>
					</xsl:when>
					<xsl:when test="w:p/w:pPr/w:jc">
						<xsl:attribute name="align"><xsl:value-of select="w:p/w:pPr/w:jc/@w:val"/></xsl:attribute>
					</xsl:when>
					<xsl:otherwise>
						<xsl:attribute name="align">left</xsl:attribute>			
					</xsl:otherwise>
				</xsl:choose>
				<xsl:apply-templates/>
			</xsl:element><br/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:for-each>
	<xsl:choose>
		<xsl:when test="w:tc[$Col]/w:tcPr/w:gridSpan/@w:val">
			<xsl:call-template name="myCells">
				<xsl:with-param name="n" select="($n + w:tc[$Col]/w:tcPr/w:gridSpan/@w:val - 1) + 1" />
				<xsl:with-param name="Col" select="$Col + 1" />
			</xsl:call-template>
		</xsl:when>
		<xsl:otherwise>
			<xsl:call-template name="myCells">
				<xsl:with-param name="n" select="$n + 1" />
				<xsl:with-param name="Col" select="$Col + 1" />
			</xsl:call-template>
		</xsl:otherwise>
	</xsl:choose>
	</xsl:if>
</xsl:template>

</xsl:stylesheet>