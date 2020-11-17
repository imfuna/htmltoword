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
  <xsl:include href="./links.xslt"/>
  <xsl:include href="./images.xslt"/>
  <xsl:template match="/">
    <xsl:apply-templates />
  </xsl:template>

  <xsl:template match="head" />

  <xsl:template match="body">
    <xsl:comment>
      KNOWN BUGS:
      div
        h2
        div
          textnode (WONT BE WRAPPED IN A W:P)
          div
            table
            span
              text
    </xsl:comment>
    <xsl:apply-templates select="*[name() != 'header' and name() != 'footer']"/>
  </xsl:template>
  <xsl:template match="section[not(contains(concat(' ', @class, ' '), ' no-banners ')) and contains(concat(' ', @class, ' '), ' page-1 ')]">
    <xsl:apply-templates />
      <w:p>
        <w:pPr>
          <w:sectPr>
            <w:pgNumType w:start="1"/>
            <w:headerReference w:type="default" r:id="rId8"/>
            <w:footerReference w:type="default" r:id="rId9"/>
            <xsl:choose>
            <xsl:when test="contains(concat(' ', @class, ' '), ' portrait ') and contains(concat(' ', @class, ' '), ' A4 ') ">
              <w:pgSz w:w="11907" w:h="16839" />
            </xsl:when>
              <xsl:when test="contains(concat(' ', @class, ' '), ' portrait ') and contains(concat(' ', @class, ' '), ' US-Letter ') ">
                <w:pgSz w:w="12240" w:h="15840"/>
              </xsl:when>
              <xsl:when test="contains(concat(' ', @class, ' '), ' landscape ') and contains(concat(' ', @class, ' '), ' A4 ') ">
                <w:pgSz w:w="16839" w:h="11907" w:orient="landscape"/>
              </xsl:when>
              <xsl:when test="contains(concat(' ', @class, ' '), ' landscape ') and contains(concat(' ', @class, ' '), ' US-Letter ') ">
                <w:pgSz w:w="15842" w:h="12242" w:orient="landscape"/>
              </xsl:when>
            </xsl:choose>
          </w:sectPr>
        </w:pPr>
      </w:p>
  </xsl:template>

  <xsl:template match="section[not(contains(concat(' ', @class, ' '), ' no-banners ')) and not(contains(concat(' ', @class, ' '), ' page-1 '))]">
    <xsl:apply-templates />
    <xsl:if test="position()!=last()">
      <w:p>
        <w:pPr>
          <w:sectPr>
            <w:headerReference w:type="default" r:id="rId8"/>
            <w:footerReference w:type="default" r:id="rId9"/>
            <xsl:choose>
              <xsl:when test="contains(concat(' ', @class, ' '), ' portrait ') and contains(concat(' ', @class, ' '), ' A4 ') ">
                <w:pgSz w:w="11907" w:h="16839" />
              </xsl:when>
              <xsl:when test="contains(concat(' ', @class, ' '), ' portrait ') and contains(concat(' ', @class, ' '), ' US-Letter ') ">
                <w:pgSz w:w="12240" w:h="15840"/>
              </xsl:when>
              <xsl:when test="contains(concat(' ', @class, ' '), ' landscape ') and contains(concat(' ', @class, ' '), ' A4 ') ">
                <w:pgSz w:w="16839" w:h="11907" w:orient="landscape"/>
              </xsl:when>
              <xsl:when test="contains(concat(' ', @class, ' '), ' landscape ') and contains(concat(' ', @class, ' '), ' US-Letter ') ">
                <w:pgSz w:w="15842" w:h="12242" w:orient="landscape"/>
              </xsl:when>
            </xsl:choose>
          </w:sectPr>
        </w:pPr>
      </w:p>
    </xsl:if>
  </xsl:template>

  <xsl:template match="section[contains(concat(' ', @class, ' '), ' no-banners ')]">
    <xsl:apply-templates />
    <w:p>
      <w:pPr>
        <w:sectPr>
          <w:hdr w:type="odd" >
            <w:p>
              <w:pPr>
                <w:pStyle w:val="Header"/>
              </w:pPr>
              <w:r>
                <w:t></w:t>
              </w:r>
            </w:p>
          </w:hdr>
          <w:ftr w:type="odd">
            <w:p>
              <w:pPr>
                <w:pStyle w:val="Footer"/>
              </w:pPr>
              <w:r>
                <w:t></w:t>
              </w:r>
            </w:p>
          </w:ftr>
          <xsl:choose>
            <xsl:when test="contains(concat(' ', @class, ' '), ' portrait ') and contains(concat(' ', @class, ' '), ' A4 ') ">
              <w:pgSz w:w="11907" w:h="16839" />
            </xsl:when>
            <xsl:when test="contains(concat(' ', @class, ' '), ' portrait ') and contains(concat(' ', @class, ' '), ' US-Letter ') ">
              <w:pgSz w:w="12240" w:h="15840"/>
            </xsl:when>
            <xsl:when test="contains(concat(' ', @class, ' '), ' landscape ') and contains(concat(' ', @class, ' '), ' A4 ') ">
              <w:pgSz w:w="16839" w:h="11907" w:orient="landscape"/>
            </xsl:when>
            <xsl:when test="contains(concat(' ', @class, ' '), ' landscape ') and contains(concat(' ', @class, ' '), ' US-Letter ') ">
              <w:pgSz w:w="15842" w:h="12242" w:orient="landscape"/>
            </xsl:when>
          </xsl:choose>
        </w:sectPr>
      </w:pPr>
    </w:p>
  </xsl:template>

  <!-- think this is looking for nodes with no children and just replaces them with running text -->
  <xsl:template match="body/*[not(*)]">
    <w:p>
      <xsl:call-template name="text-alignment" />
      <w:r>
        <w:t xml:space="preserve"><xsl:value-of select="."/></w:t>
      </w:r>
    </w:p>
  </xsl:template>

  <xsl:template match="br[not(ancestor::p) and not(ancestor::div) and not(ancestor::td|ancestor::li) or
                          (preceding-sibling::div or following-sibling::div or preceding-sibling::p or following-sibling::p)]">
    <w:p>
      <w:r></w:r>
    </w:p>
  </xsl:template>

  <xsl:template match="br[(ancestor::li or ancestor::td) and
                          (preceding-sibling::div or following-sibling::div or preceding-sibling::p or following-sibling::p)]">
    <w:r>
      <w:br />
    </w:r>
  </xsl:template>

  <xsl:template match="br">
    <w:r>
      <w:br />
    </w:r>
  </xsl:template>

  <xsl:template match="pre">
    <w:p>
      <xsl:apply-templates />
    </w:p>
  </xsl:template>

  <xsl:template match="div[not(ancestor::li) and not(ancestor::td) and not(ancestor::th) and not(ancestor::p) and not(ancestor::dl) and not(ancestor::header) and not(descendant::dl) and not(descendant::div) and not(descendant::p) and not(descendant::h1) and not(descendant::h2) and not(descendant::h3) and not(descendant::h4) and not(descendant::h5) and not(descendant::h6) and not(descendant::table) and not(descendant::li) and not (descendant::pre)]">
    <xsl:comment>Divs should create a p if nothing above them has and nothing below them will XXX version 1.1.8</xsl:comment>
    <w:p>
      <xsl:call-template name="text-alignment" />
      <xsl:apply-templates />
    </w:p>
  </xsl:template>

  <xsl:template match="div">
    <xsl:apply-templates />
  </xsl:template>

  <!-- TODO: make this prettier. Headings shouldn't enter in template from L51 -->
  <xsl:template match="body/h1|body/h2|body/h3|body/h4|body/h5|body/h6|h1|h2|h3|h4|h5|h6">
    <xsl:variable name="length" select="string-length(name(.))"/>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading{substring(name(.),$length)}"/>
      </w:pPr>
      <w:r>
        <w:t xml:space="preserve"><xsl:value-of select="."/></w:t>
      </w:r>
    </w:p>
  </xsl:template>

  <xsl:template match="p[not(ancestor::li)]">
    <w:p>
      <xsl:call-template name="text-alignment" />
      <xsl:apply-templates />
    </w:p>
  </xsl:template>

  <xsl:template match="ol|ul">
    <xsl:param name="global_level" select="count(preceding::ol[not(ancestor::ol or ancestor::ul)]) + count(preceding::ul[not(ancestor::ol or ancestor::ul)]) + 1"/>
    <xsl:apply-templates>
      <xsl:with-param name="global_level" select="$global_level" />
    </xsl:apply-templates>
  </xsl:template>

  <xsl:template name="listItem" match="li">
    <xsl:param name="global_level" />
    <xsl:param name="preceding-siblings" select="0"/>
    <xsl:for-each select="node()">
      <xsl:choose>
        <xsl:when test="self::br">
          <w:p>
            <w:pPr>
              <w:pStyle w:val="ListParagraph"></w:pStyle>
              <w:numPr>
                <w:ilvl w:val="0"/>
                <w:numId w:val="0"/>
              </w:numPr>
            </w:pPr>
            <w:r></w:r>
          </w:p>
        </xsl:when>
        <xsl:when test="self::ol|self::ul">
          <xsl:apply-templates>
            <xsl:with-param name="global_level" select="$global_level" />
          </xsl:apply-templates>
        </xsl:when>
        <xsl:when test="not(descendant::li)"> <!-- simpler div, p, headings, etc... -->
          <xsl:variable name="ilvl" select="count(ancestor::ol) + count(ancestor::ul) - 1"></xsl:variable>
          <xsl:choose>
            <xsl:when test="$preceding-siblings + count(preceding-sibling::*) > 0">
              <w:p>
                <w:pPr>
                  <w:pStyle w:val="ListParagraph"></w:pStyle>
                  <w:numPr>
                    <w:ilvl w:val="0"/>
                    <w:numId w:val="0"/>
                  </w:numPr>
                  <w:ind w:left="{720 * ($ilvl + 1)}"/>
                </w:pPr>
                <xsl:apply-templates/>
              </w:p>
            </xsl:when>
            <xsl:otherwise>
              <w:p>
                <w:pPr>
                  <w:pStyle w:val="ListParagraph"></w:pStyle>
                  <w:numPr>
                    <w:ilvl w:val="{$ilvl}"/>
                    <w:numId w:val="{$global_level}"/>
                  </w:numPr>
                </w:pPr>
                <xsl:apply-templates/>
              </w:p>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise> <!-- mixed things div having list and stuff content... -->
          <xsl:call-template name="listItem">
            <xsl:with-param name="global_level" select="$global_level" />
            <xsl:with-param name="preceding-siblings" select="$preceding-siblings + count(preceding-sibling::*)" />
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="dl">
    <xsl:apply-templates/>
  </xsl:template>

  <xsl:template match="dt">
    <w:p>
      <xsl:apply-templates/>
    </w:p>
  </xsl:template>

  <xsl:template match="dd">
    <w:p>
      <w:pPr>
        <w:ind w:left="720"/>
      </w:pPr>
      <xsl:apply-templates/>
    </w:p>
  </xsl:template>

  <xsl:template match="span[not(ancestor::td) and not(ancestor::li) and (preceding-sibling::h1 or preceding-sibling::h2 or preceding-sibling::h3 or preceding-sibling::h4 or preceding-sibling::h5 or preceding-sibling::h6 or preceding-sibling::table or preceding-sibling::p or preceding-sibling::ol or preceding-sibling::ul or preceding-sibling::div or following-sibling::h1 or following-sibling::h2 or following-sibling::h3 or following-sibling::h4 or following-sibling::h5 or following-sibling::h6 or following-sibling::table or following-sibling::p or following-sibling::ol or following-sibling::ul or following-sibling::div)]
    |a[not(ancestor::td) and not(ancestor::li) and (preceding-sibling::h1 or preceding-sibling::h2 or preceding-sibling::h3 or preceding-sibling::h4 or preceding-sibling::h5 or preceding-sibling::h6 or preceding-sibling::table or preceding-sibling::p or preceding-sibling::ol or preceding-sibling::ul or preceding-sibling::div or following-sibling::h1 or following-sibling::h2 or following-sibling::h3 or following-sibling::h4 or following-sibling::h5 or following-sibling::h6 or following-sibling::table or following-sibling::p or following-sibling::ol or following-sibling::ul or following-sibling::div)]
    |small[not(ancestor::td) and not(ancestor::li) and (preceding-sibling::h1 or preceding-sibling::h2 or preceding-sibling::h3 or preceding-sibling::h4 or preceding-sibling::h5 or preceding-sibling::h6 or preceding-sibling::table or preceding-sibling::p or preceding-sibling::ol or preceding-sibling::ul or preceding-sibling::div or following-sibling::h1 or following-sibling::h2 or following-sibling::h3 or following-sibling::h4 or following-sibling::h5 or following-sibling::h6 or following-sibling::table or following-sibling::p or following-sibling::ol or following-sibling::ul or following-sibling::div)]
    |strong[not(ancestor::td) and not(ancestor::li) and (preceding-sibling::h1 or preceding-sibling::h2 or preceding-sibling::h3 or preceding-sibling::h4 or preceding-sibling::h5 or preceding-sibling::h6 or preceding-sibling::table or preceding-sibling::p or preceding-sibling::ol or preceding-sibling::ul or preceding-sibling::div or following-sibling::h1 or following-sibling::h2 or following-sibling::h3 or following-sibling::h4 or following-sibling::h5 or following-sibling::h6 or following-sibling::table or following-sibling::p or following-sibling::ol or following-sibling::ul or following-sibling::div)]
    |em[not(ancestor::td) and not(ancestor::li) and (preceding-sibling::h1 or preceding-sibling::h2 or preceding-sibling::h3 or preceding-sibling::h4 or preceding-sibling::h5 or preceding-sibling::h6 or preceding-sibling::table or preceding-sibling::p or preceding-sibling::ol or preceding-sibling::ul or preceding-sibling::div or following-sibling::h1 or following-sibling::h2 or following-sibling::h3 or following-sibling::h4 or following-sibling::h5 or following-sibling::h6 or following-sibling::table or following-sibling::p or following-sibling::ol or following-sibling::ul or following-sibling::div)]
    |i[not(ancestor::td) and not(ancestor::li) and (preceding-sibling::h1 or preceding-sibling::h2 or preceding-sibling::h3 or preceding-sibling::h4 or preceding-sibling::h5 or preceding-sibling::h6 or preceding-sibling::table or preceding-sibling::p or preceding-sibling::ol or preceding-sibling::ul or preceding-sibling::div or following-sibling::h1 or following-sibling::h2 or following-sibling::h3 or following-sibling::h4 or following-sibling::h5 or following-sibling::h6 or following-sibling::table or following-sibling::p or following-sibling::ol or following-sibling::ul or following-sibling::div)]
    |b[not(ancestor::td) and not(ancestor::li) and (preceding-sibling::h1 or preceding-sibling::h2 or preceding-sibling::h3 or preceding-sibling::h4 or preceding-sibling::h5 or preceding-sibling::h6 or preceding-sibling::table or preceding-sibling::p or preceding-sibling::ol or preceding-sibling::ul or preceding-sibling::div or following-sibling::h1 or following-sibling::h2 or following-sibling::h3 or following-sibling::h4 or following-sibling::h5 or following-sibling::h6 or following-sibling::table or following-sibling::p or following-sibling::ol or following-sibling::ul or following-sibling::div)]
    |u[not(ancestor::td) and not(ancestor::li) and (preceding-sibling::h1 or preceding-sibling::h2 or preceding-sibling::h3 or preceding-sibling::h4 or preceding-sibling::h5 or preceding-sibling::h6 or preceding-sibling::table or preceding-sibling::p or preceding-sibling::ol or preceding-sibling::ul or preceding-sibling::div or following-sibling::h1 or following-sibling::h2 or following-sibling::h3 or following-sibling::h4 or following-sibling::h5 or following-sibling::h6 or following-sibling::table or following-sibling::p or following-sibling::ol or following-sibling::ul or following-sibling::div)]">
    <xsl:comment>
        In the following situation:

        div
          h2
          span
            textnode
            span
              textnode
          p

        The div template will not create a w:p because the div contains a h2. Therefore we need to wrap the inline elements span|a|small in a p here.
      </xsl:comment>
    <w:p>
      <xsl:choose>
        <xsl:when test="self::a[starts-with(@href, 'http://') or starts-with(@href, 'https://')]">
          <xsl:call-template name="link" />
        </xsl:when>
        <xsl:when test="self::img">
          <xsl:comment>
            This template adds images.
          </xsl:comment>
          <xsl:call-template name="image"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:apply-templates />
        </xsl:otherwise>
      </xsl:choose>
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

  <xsl:template match="span[contains(concat(' ', @class, ' '), ' h ')]">
    <xsl:comment>
        This template adds MS Word highlighting ability.
      </xsl:comment>
    <xsl:variable name="color">
      <xsl:choose>
        <xsl:when test="./@data-style='pink'">magenta</xsl:when>
        <xsl:when test="./@data-style='blue'">cyan</xsl:when>
        <xsl:when test="./@data-style='orange'">darkYellow</xsl:when>
        <xsl:otherwise><xsl:value-of select="./@data-style"/></xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="preceding-sibling::h1 or preceding-sibling::h2 or preceding-sibling::h3 or preceding-sibling::h4 or preceding-sibling::h5 or preceding-sibling::h6 or preceding-sibling::table or preceding-sibling::p or preceding-sibling::ol or preceding-sibling::ul or preceding-sibling::div or following-sibling::h1 or following-sibling::h2 or following-sibling::h3 or following-sibling::h4 or following-sibling::h5 or following-sibling::h6 or following-sibling::table or following-sibling::p or following-sibling::ol or following-sibling::ul or following-sibling::div">
        <w:p>
          <w:r>
            <w:rPr>
              <w:highlight w:val="{$color}"/>
            </w:rPr>
            <w:t xml:space="preserve"><xsl:value-of select="."/></w:t>
          </w:r>
        </w:p>
      </xsl:when>
      <xsl:otherwise>
        <w:r>
          <w:rPr>
            <w:highlight w:val="{$color}"/>
          </w:rPr>
          <w:t xml:space="preserve"><xsl:value-of select="."/></w:t>
        </w:r>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="div[contains(concat(' ', @class, ' '), ' -page-break ')]">
    <w:p>
      <w:r>
        <w:br w:type="page" />
      </w:r>
    </w:p>
    <xsl:apply-templates />
  </xsl:template>

  <xsl:template match="div[contains(concat(' ', @class, ' '), ' toc ')]">
    <w:sdt>
    <w:sdtPr>
      <w:id w:val="-493258456" />
      <w:docPartObj>
        <w:docPartGallery w:val="Table of Contents" />
        <w:docPartUnique />
      </w:docPartObj>
    </w:sdtPr>
    <w:sdtEndPr>
      <w:rPr>
        <w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi" />
        <w:b />
        <w:bCs />
        <w:noProof />
        <w:color w:val="auto" />
        <w:sz w:val="22" />
        <w:szCs w:val="22" />
      </w:rPr>
    </w:sdtEndPr>
    <w:sdtContent>
      <w:p w:rsidR="00095C65" w:rsidRDefault="00095C65">
      <w:pPr>
        <w:pStyle w:val="TOCHeading" />
        <w:jc w:val="center" />
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b />
          <w:color w:val="2E74B5" w:themeColor="accent1" w:themeShade="BF" />
          <w:sz w:val="28" />
          <w:szCs w:val="28" />
        </w:rPr>
        <w:t>Table of Contents</w:t>
      </w:r>
    </w:p>
    <w:p w:rsidR="00095C65" w:rsidRDefault="00095C65">
    <w:r>
      <w:rPr>
        <w:b />
        <w:bCs />
        <w:noProof />
      </w:rPr>
      <w:fldChar w:fldCharType="begin" />
    </w:r>
    <w:r>
      <w:rPr>
        <w:b />
        <w:bCs />
        <w:noProof />
      </w:rPr>
      <w:instrText xml:space="preserve"> TOC \o "1-2" \h \z \u  \t "Heading 1,1,Heading 2,2"</w:instrText>
  </w:r>
    <w:r>
      <w:rPr>
        <w:b />
        <w:bCs />
        <w:noProof />
      </w:rPr>
      <w:fldChar w:fldCharType="separate" />
    </w:r>
    <w:r>
      <w:rPr>
        <w:noProof />
      </w:rPr>
      <w:t>No table of contents entries found.</w:t>
    </w:r>
    <w:r>
      <w:rPr>
        <w:b />
        <w:bCs />
        <w:noProof />
      </w:rPr>
      <w:fldChar w:fldCharType="end" />
    </w:r>
  </w:p>
</w:sdtContent>
    </w:sdt>
  </xsl:template>

  <xsl:template match="details" />

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

  <xsl:template match="*">
    <xsl:apply-templates/>
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
