<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
                version="1.0">
  <xsl:output method="xml" encoding="utf-8" omit-xml-declaration="yes" indent="yes" />

  <xsl:include href="./image_functions.xslt"/>
  <xsl:include href="./functions.xslt"/>

  <xsl:template match="/*">
    <xsl:apply-templates select='/html/body/footer'/>
  </xsl:template>

  <xsl:template match="footer">
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships>
      <xsl:apply-templates select="*"/>
    </Relationships>
  </xsl:template>

  <!--
    Images can either be embedded in the document or referenced via the external url.
    By default images are embedded and a copy will be stored in the final zip folder.
    If you want your images to be referenced externally (don't store the image files) then add  data-external="true" to your image tag(s).
    If you are embedding them images need a filename. You can either define one by passing the name and extension to data-filename or rely on the defaul which is the last part of the source url.
    **Word files can be corrupted if an extension isn't present so if you are relying on the source url ensure that is has an extension.**
  -->
  <xsl:template match="img" priority="0">
    <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image">
      <xsl:choose>
        <xsl:when test="@data-external = 'true'">
          <xsl:attribute name="Target"><xsl:value-of select="@src"/></xsl:attribute>
          <xsl:attribute name="TargetMode">External</xsl:attribute>
        </xsl:when>

        <xsl:otherwise>
          <xsl:attribute name="Target">media/image<xsl:value-of select="count(preceding::img)+1"/>.<xsl:call-template name="image-extension">
              <xsl:with-param name="source" select="@src"/>
              <xsl:with-param name="data-filename" select="@data-filename"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:attribute name="Id"><xsl:call-template name="relationship-id"/></xsl:attribute>
    </Relationship>
  </xsl:template>


  <xsl:template match="text()|@*"/>
</xsl:stylesheet>
