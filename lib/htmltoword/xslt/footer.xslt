<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:o="urn:schemas-microsoft-com:office:office"
                xmlns:v="urn:schemas-microsoft-com:vml"
                xmlns:WX="http://schemas.microsoft.com/office/word/2003/auxHint"
                xmlns:aml="http://schemas.microsoft.com/aml/2001/core"
                xmlns:w10="urn:schemas-microsoft-com:office:word"
                xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                xmlns:ext="http://www.xmllab.net/wordml2html/ext"
                xmlns:java="http://xml.apache.org/xalan/java"
                xmlns:str="http://exslt.org/strings"
                xmlns:func="http://exslt.org/functions"
                xmlns:fn="http://www.w3.org/2005/xpath-functions"
                version="1.0"
                exclude-result-prefixes="java msxsl ext w o v WX aml w10"
                extension-element-prefixes="func">
  <xsl:output method="xml" encoding="utf-8" omit-xml-declaration="yes" indent="yes" />
  <xsl:include href="./functions.xslt"/>
  <xsl:include href="./tables.xslt"/>
  <xsl:include href="./images.xslt"/>
  <!--<xsl:template match="/">-->
    <!--<xsl:apply-templates />-->
  <!--</xsl:template>-->

  <xsl:template match="/*">
    <xsl:apply-templates select='/html/body/footer'/>
  </xsl:template>

  <xsl:template match="footer">
    <w:ftr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14">
      <xsl:apply-templates select="table"/>
      <w:p w14:paraId="042D5011" w14:textId="784386AD" w:rsidR="00FD1288" w:rsidRDefault="00FD1288" w:rsidP="004D351C">
        <w:pPr><w:pStyle w:val="FooterStyle"/></w:pPr>
      </w:p>
    </w:ftr>
  </xsl:template>

  <xsl:template match="p[not(ancestor::li)]">
    <w:p>
      <xsl:call-template name="text-alignment" />
      <xsl:apply-templates />
    </w:p>
  </xsl:template>

  <xsl:template match="text()[not(parent::li) and not(parent::td) and not(parent::pre) and (preceding-sibling::h1 or preceding-sibling::h2 or preceding-sibling::h3 or preceding-sibling::h4 or preceding-sibling::h5 or preceding-sibling::h6 or preceding-sibling::table or preceding-sibling::p or preceding-sibling::ol or preceding-sibling::ul or preceding-sibling::div or following-sibling::h1 or following-sibling::h2 or following-sibling::h3 or following-sibling::h4 or following-sibling::h5 or following-sibling::h6 or following-sibling::table or following-sibling::p or following-sibling::ol or following-sibling::ul or following-sibling::div)]">
    <xsl:comment>
      In the following situation:

      div
      h2
      textnode
      p

      The div template will not create a w:p because the div contains a h2. Therefore we need to wrap the textnode in a p here.
    </xsl:comment>
    <w:p>
      <w:r>
        <w:t xml:space="preserve"><xsl:value-of select="."/></w:t>
      </w:r>
    </w:p>
  </xsl:template>

  <xsl:template match="span[contains(concat(' ', @class, ' '), ' page_number ')]">
    <xsl:comment>page number</xsl:comment>
    <w:pPr>
      <w:pStyle w:val="PageNumberStyle"/>
    </w:pPr>
    <w:r>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
      <w:instrText xml:space="preserve">PAGE
      </w:instrText>
    </w:r>
    <w:r>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:noProof/>
      </w:rPr>
      <w:t>1</w:t>
    </w:r>
    <w:r>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
  </xsl:template>

  <xsl:template match="text()">
    <xsl:if test="string-length(.) > 0">
      <w:r>
        <xsl:if test="ancestor::i">
          <w:rPr>
            <w:i />
          </w:rPr>
        </xsl:if>
        <xsl:if test="ancestor::b">
          <w:rPr>
            <w:b />
          </w:rPr>
        </xsl:if>
        <xsl:if test="ancestor::u">
          <w:rPr>
            <w:u w:val="single"/>
          </w:rPr>
        </xsl:if>
        <xsl:if test="ancestor::s">
          <w:rPr>
            <w:strike w:val="true"/>
          </w:rPr>
        </xsl:if>
        <xsl:if test="ancestor::sub">
          <w:rPr>
            <w:vertAlign w:val="subscript"/>
          </w:rPr>
        </xsl:if>
        <xsl:if test="ancestor::sup">
          <w:rPr>
            <w:vertAlign w:val="superscript"/>
          </w:rPr>
        </xsl:if>
        <w:t xml:space="preserve"><xsl:value-of select="."/></w:t>
      </w:r>
    </xsl:if>
  </xsl:template>

  <xsl:template name="text-alignment">
    <xsl:param name="class" select="@class" />
    <xsl:param name="style" select="@style" />
    <xsl:param name="element-style" select="@data-style" />
    <xsl:comment>Class</xsl:comment>
    <xsl:comment><xsl:copy-of select="$class"/></xsl:comment>
    <xsl:comment>Parent Class</xsl:comment>
    <xsl:comment><xsl:copy-of select="../@class"/></xsl:comment>

    <xsl:variable name="alignment">
      <xsl:choose>
        <xsl:when test="contains(concat(' ', $class, ' '), ' center ') or contains(translate(normalize-space($style),' ',''), 'text-align:center')">center</xsl:when>
        <xsl:when test="contains(concat(' ', $class, ' '), ' right ') or contains(translate(normalize-space($style),' ',''), 'text-align:right')">right</xsl:when>
        <xsl:when test="contains(concat(' ', $class, ' '), ' left ') or contains(translate(normalize-space($style),' ',''), 'text-align:left')">left</xsl:when>
        <xsl:when test="contains(concat(' ', $class, ' '), ' justify ') or contains(translate(normalize-space($style),' ',''), 'text-align:justify')">both</xsl:when>
        <xsl:otherwise></xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="needs-element-style" >
      <xsl:choose>
        <xsl:when test="contains(concat(' ', $class, ' '), ' h-style ') and string-length(normalize-space($element-style)) > 0">yes</xsl:when>
      </xsl:choose>
    </xsl:variable>
    <xsl:if test="string-length(normalize-space($alignment)) > 0 or string-length(normalize-space($needs-element-style)) > 0">
      <w:pPr>
        <xsl:if test="string-length(normalize-space($alignment)) > 0">
          <w:jc w:val="{$alignment}"/>
        </xsl:if>
        <xsl:if test="string-length(normalize-space($needs-element-style)) > 0">
          <w:pStyle w:val="{$element-style}"/>
        </xsl:if>
      </w:pPr>
    </xsl:if>
  </xsl:template>

</xsl:stylesheet>