package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface PageSetup extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   */

  @DISPID(148)
  @PropGet
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   */

  @DISPID(149)
  @PropGet
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   */

  @DISPID(150)
  @PropGet
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "BlackAndWhite"
   * </p>
   */

  @DISPID(1009)
  @PropGet
  boolean getBlackAndWhite();


  /**
   * <p>
   * Setter method for the COM property "BlackAndWhite"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1009)
  @PropPut
  void setBlackAndWhite(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "BottomMargin"
   * </p>
   */

  @DISPID(1002)
  @PropGet
  double getBottomMargin();


  /**
   * <p>
   * Setter method for the COM property "BottomMargin"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(1002)
  @PropPut
  void setBottomMargin(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "CenterFooter"
   * </p>
   */

  @DISPID(1010)
  @PropGet
  java.lang.String getCenterFooter();


  /**
   * <p>
   * Setter method for the COM property "CenterFooter"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1010)
  @PropPut
  void setCenterFooter(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "CenterHeader"
   * </p>
   */

  @DISPID(1011)
  @PropGet
  java.lang.String getCenterHeader();


  /**
   * <p>
   * Setter method for the COM property "CenterHeader"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1011)
  @PropPut
  void setCenterHeader(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "CenterHorizontally"
   * </p>
   */

  @DISPID(1005)
  @PropGet
  boolean getCenterHorizontally();


  /**
   * <p>
   * Setter method for the COM property "CenterHorizontally"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1005)
  @PropPut
  void setCenterHorizontally(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "CenterVertically"
   * </p>
   */

  @DISPID(1006)
  @PropGet
  boolean getCenterVertically();


  /**
   * <p>
   * Setter method for the COM property "CenterVertically"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1006)
  @PropPut
  void setCenterVertically(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ChartSize"
   * </p>
   */

  @DISPID(1012)
  @PropGet
  com.exceljava.com4j.excel.XlObjectSize getChartSize();


  /**
   * <p>
   * Setter method for the COM property "ChartSize"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlObjectSize parameter.
   */

  @DISPID(1012)
  @PropPut
  void setChartSize(
    com.exceljava.com4j.excel.XlObjectSize rhs);


  /**
   * <p>
   * Getter method for the COM property "Draft"
   * </p>
   */

  @DISPID(1020)
  @PropGet
  boolean getDraft();


  /**
   * <p>
   * Setter method for the COM property "Draft"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1020)
  @PropPut
  void setDraft(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "FirstPageNumber"
   * </p>
   */

  @DISPID(1008)
  @PropGet
  int getFirstPageNumber();


  /**
   * <p>
   * Setter method for the COM property "FirstPageNumber"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(1008)
  @PropPut
  void setFirstPageNumber(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "FitToPagesTall"
   * </p>
   */

  @DISPID(1013)
  @PropGet
  java.lang.Object getFitToPagesTall();


  /**
   * <p>
   * Setter method for the COM property "FitToPagesTall"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1013)
  @PropPut
  void setFitToPagesTall(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "FitToPagesWide"
   * </p>
   */

  @DISPID(1014)
  @PropGet
  java.lang.Object getFitToPagesWide();


  /**
   * <p>
   * Setter method for the COM property "FitToPagesWide"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1014)
  @PropPut
  void setFitToPagesWide(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "FooterMargin"
   * </p>
   */

  @DISPID(1015)
  @PropGet
  double getFooterMargin();


  /**
   * <p>
   * Setter method for the COM property "FooterMargin"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(1015)
  @PropPut
  void setFooterMargin(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "HeaderMargin"
   * </p>
   */

  @DISPID(1016)
  @PropGet
  double getHeaderMargin();


  /**
   * <p>
   * Setter method for the COM property "HeaderMargin"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(1016)
  @PropPut
  void setHeaderMargin(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "LeftFooter"
   * </p>
   */

  @DISPID(1017)
  @PropGet
  java.lang.String getLeftFooter();


  /**
   * <p>
   * Setter method for the COM property "LeftFooter"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1017)
  @PropPut
  void setLeftFooter(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "LeftHeader"
   * </p>
   */

  @DISPID(1018)
  @PropGet
  java.lang.String getLeftHeader();


  /**
   * <p>
   * Setter method for the COM property "LeftHeader"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1018)
  @PropPut
  void setLeftHeader(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "LeftMargin"
   * </p>
   */

  @DISPID(999)
  @PropGet
  double getLeftMargin();


  /**
   * <p>
   * Setter method for the COM property "LeftMargin"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(999)
  @PropPut
  void setLeftMargin(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Order"
   * </p>
   */

  @DISPID(192)
  @PropGet
  com.exceljava.com4j.excel.XlOrder getOrder();


  /**
   * <p>
   * Setter method for the COM property "Order"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlOrder parameter.
   */

  @DISPID(192)
  @PropPut
  void setOrder(
    com.exceljava.com4j.excel.XlOrder rhs);


  /**
   * <p>
   * Getter method for the COM property "Orientation"
   * </p>
   */

  @DISPID(134)
  @PropGet
  com.exceljava.com4j.excel.XlPageOrientation getOrientation();


  /**
   * <p>
   * Setter method for the COM property "Orientation"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlPageOrientation parameter.
   */

  @DISPID(134)
  @PropPut
  void setOrientation(
    com.exceljava.com4j.excel.XlPageOrientation rhs);


  /**
   * <p>
   * Getter method for the COM property "PaperSize"
   * </p>
   */

  @DISPID(1007)
  @PropGet
  com.exceljava.com4j.excel.XlPaperSize getPaperSize();


  /**
   * <p>
   * Setter method for the COM property "PaperSize"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlPaperSize parameter.
   */

  @DISPID(1007)
  @PropPut
  void setPaperSize(
    com.exceljava.com4j.excel.XlPaperSize rhs);


  /**
   * <p>
   * Getter method for the COM property "PrintArea"
   * </p>
   */

  @DISPID(1019)
  @PropGet
  java.lang.String getPrintArea();


  /**
   * <p>
   * Setter method for the COM property "PrintArea"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1019)
  @PropPut
  void setPrintArea(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "PrintGridlines"
   * </p>
   */

  @DISPID(1004)
  @PropGet
  boolean getPrintGridlines();


  /**
   * <p>
   * Setter method for the COM property "PrintGridlines"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1004)
  @PropPut
  void setPrintGridlines(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "PrintHeadings"
   * </p>
   */

  @DISPID(1003)
  @PropGet
  boolean getPrintHeadings();


  /**
   * <p>
   * Setter method for the COM property "PrintHeadings"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1003)
  @PropPut
  void setPrintHeadings(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "PrintNotes"
   * </p>
   */

  @DISPID(1021)
  @PropGet
  boolean getPrintNotes();


  /**
   * <p>
   * Setter method for the COM property "PrintNotes"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1021)
  @PropPut
  void setPrintNotes(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "PrintQuality"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getPrintQuality(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1022)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object getPrintQuality();

  /**
   * <p>
   * Getter method for the COM property "PrintQuality"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1022)
  @PropGet
  java.lang.Object getPrintQuality(
    @Optional java.lang.Object index);


  /**
   * <p>
   * Setter method for the COM property "PrintQuality"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setPrintQuality(com4j.Variant.getMissing(), rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1022)
  @PropPut
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setPrintQuality(
    java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "PrintQuality"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1022)
  @PropPut
  void setPrintQuality(
    @Optional java.lang.Object index,
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "PrintTitleColumns"
   * </p>
   */

  @DISPID(1023)
  @PropGet
  java.lang.String getPrintTitleColumns();


  /**
   * <p>
   * Setter method for the COM property "PrintTitleColumns"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1023)
  @PropPut
  void setPrintTitleColumns(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "PrintTitleRows"
   * </p>
   */

  @DISPID(1024)
  @PropGet
  java.lang.String getPrintTitleRows();


  /**
   * <p>
   * Setter method for the COM property "PrintTitleRows"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1024)
  @PropPut
  void setPrintTitleRows(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "RightFooter"
   * </p>
   */

  @DISPID(1025)
  @PropGet
  java.lang.String getRightFooter();


  /**
   * <p>
   * Setter method for the COM property "RightFooter"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1025)
  @PropPut
  void setRightFooter(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "RightHeader"
   * </p>
   */

  @DISPID(1026)
  @PropGet
  java.lang.String getRightHeader();


  /**
   * <p>
   * Setter method for the COM property "RightHeader"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1026)
  @PropPut
  void setRightHeader(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "RightMargin"
   * </p>
   */

  @DISPID(1000)
  @PropGet
  double getRightMargin();


  /**
   * <p>
   * Setter method for the COM property "RightMargin"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(1000)
  @PropPut
  void setRightMargin(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "TopMargin"
   * </p>
   */

  @DISPID(1001)
  @PropGet
  double getTopMargin();


  /**
   * <p>
   * Setter method for the COM property "TopMargin"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(1001)
  @PropPut
  void setTopMargin(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Zoom"
   * </p>
   */

  @DISPID(663)
  @PropGet
  java.lang.Object getZoom();


  /**
   * <p>
   * Setter method for the COM property "Zoom"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(663)
  @PropPut
  void setZoom(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "PrintComments"
   * </p>
   */

  @DISPID(1524)
  @PropGet
  com.exceljava.com4j.excel.XlPrintLocation getPrintComments();


  /**
   * <p>
   * Setter method for the COM property "PrintComments"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlPrintLocation parameter.
   */

  @DISPID(1524)
  @PropPut
  void setPrintComments(
    com.exceljava.com4j.excel.XlPrintLocation rhs);


  /**
   * <p>
   * Getter method for the COM property "PrintErrors"
   * </p>
   */

  @DISPID(2149)
  @PropGet
  com.exceljava.com4j.excel.XlPrintErrors getPrintErrors();


  /**
   * <p>
   * Setter method for the COM property "PrintErrors"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlPrintErrors parameter.
   */

  @DISPID(2149)
  @PropPut
  void setPrintErrors(
    com.exceljava.com4j.excel.XlPrintErrors rhs);


  /**
   * <p>
   * Getter method for the COM property "CenterHeaderPicture"
   * </p>
   */

  @DISPID(2150)
  @PropGet
  com.exceljava.com4j.excel.Graphic getCenterHeaderPicture();


  /**
   * <p>
   * Getter method for the COM property "CenterFooterPicture"
   * </p>
   */

  @DISPID(2151)
  @PropGet
  com.exceljava.com4j.excel.Graphic getCenterFooterPicture();


  /**
   * <p>
   * Getter method for the COM property "LeftHeaderPicture"
   * </p>
   */

  @DISPID(2152)
  @PropGet
  com.exceljava.com4j.excel.Graphic getLeftHeaderPicture();


  /**
   * <p>
   * Getter method for the COM property "LeftFooterPicture"
   * </p>
   */

  @DISPID(2153)
  @PropGet
  com.exceljava.com4j.excel.Graphic getLeftFooterPicture();


  /**
   * <p>
   * Getter method for the COM property "RightHeaderPicture"
   * </p>
   */

  @DISPID(2154)
  @PropGet
  com.exceljava.com4j.excel.Graphic getRightHeaderPicture();


  /**
   * <p>
   * Getter method for the COM property "RightFooterPicture"
   * </p>
   */

  @DISPID(2155)
  @PropGet
  com.exceljava.com4j.excel.Graphic getRightFooterPicture();


  /**
   * <p>
   * Getter method for the COM property "OddAndEvenPagesHeaderFooter"
   * </p>
   */

  @DISPID(2600)
  @PropGet
  boolean getOddAndEvenPagesHeaderFooter();


  /**
   * <p>
   * Setter method for the COM property "OddAndEvenPagesHeaderFooter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2600)
  @PropPut
  void setOddAndEvenPagesHeaderFooter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DifferentFirstPageHeaderFooter"
   * </p>
   */

  @DISPID(2601)
  @PropGet
  boolean getDifferentFirstPageHeaderFooter();


  /**
   * <p>
   * Setter method for the COM property "DifferentFirstPageHeaderFooter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2601)
  @PropPut
  void setDifferentFirstPageHeaderFooter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ScaleWithDocHeaderFooter"
   * </p>
   */

  @DISPID(2602)
  @PropGet
  boolean getScaleWithDocHeaderFooter();


  /**
   * <p>
   * Setter method for the COM property "ScaleWithDocHeaderFooter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2602)
  @PropPut
  void setScaleWithDocHeaderFooter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "AlignMarginsHeaderFooter"
   * </p>
   */

  @DISPID(2603)
  @PropGet
  boolean getAlignMarginsHeaderFooter();


  /**
   * <p>
   * Setter method for the COM property "AlignMarginsHeaderFooter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2603)
  @PropPut
  void setAlignMarginsHeaderFooter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Pages"
   * </p>
   */

  @DISPID(2604)
  @PropGet
  com.exceljava.com4j.excel.Pages getPages();


  /**
   * <p>
   * Getter method for the COM property "EvenPage"
   * </p>
   */

  @DISPID(2605)
  @PropGet
  com.exceljava.com4j.excel.Page getEvenPage();


  /**
   * <p>
   * Getter method for the COM property "FirstPage"
   * </p>
   */

  @DISPID(2606)
  @PropGet
  com.exceljava.com4j.excel.Page getFirstPage();


  // Properties:
}
