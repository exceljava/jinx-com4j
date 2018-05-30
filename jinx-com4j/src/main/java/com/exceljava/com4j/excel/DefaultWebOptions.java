package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024448-0000-0000-C000-000000000046}")
public interface DefaultWebOptions extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Application
   */

  @DISPID(148) //= 0x94. The runtime will prefer the VTID if present
  @VTID(7)
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCreator
   */

  @DISPID(149) //= 0x95. The runtime will prefer the VTID if present
  @VTID(8)
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(150) //= 0x96. The runtime will prefer the VTID if present
  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "RelyOnCSS"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1899) //= 0x76b. The runtime will prefer the VTID if present
  @VTID(10)
  boolean getRelyOnCSS();


  /**
   * <p>
   * Setter method for the COM property "RelyOnCSS"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1899) //= 0x76b. The runtime will prefer the VTID if present
  @VTID(11)
  void setRelyOnCSS(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "SaveHiddenData"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1900) //= 0x76c. The runtime will prefer the VTID if present
  @VTID(12)
  boolean getSaveHiddenData();


  /**
   * <p>
   * Setter method for the COM property "SaveHiddenData"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1900) //= 0x76c. The runtime will prefer the VTID if present
  @VTID(13)
  void setSaveHiddenData(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "LoadPictures"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1901) //= 0x76d. The runtime will prefer the VTID if present
  @VTID(14)
  boolean getLoadPictures();


  /**
   * <p>
   * Setter method for the COM property "LoadPictures"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1901) //= 0x76d. The runtime will prefer the VTID if present
  @VTID(15)
  void setLoadPictures(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "OrganizeInFolder"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1902) //= 0x76e. The runtime will prefer the VTID if present
  @VTID(16)
  boolean getOrganizeInFolder();


  /**
   * <p>
   * Setter method for the COM property "OrganizeInFolder"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1902) //= 0x76e. The runtime will prefer the VTID if present
  @VTID(17)
  void setOrganizeInFolder(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "UpdateLinksOnSave"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1903) //= 0x76f. The runtime will prefer the VTID if present
  @VTID(18)
  boolean getUpdateLinksOnSave();


  /**
   * <p>
   * Setter method for the COM property "UpdateLinksOnSave"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1903) //= 0x76f. The runtime will prefer the VTID if present
  @VTID(19)
  void setUpdateLinksOnSave(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "UseLongFileNames"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1904) //= 0x770. The runtime will prefer the VTID if present
  @VTID(20)
  boolean getUseLongFileNames();


  /**
   * <p>
   * Setter method for the COM property "UseLongFileNames"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1904) //= 0x770. The runtime will prefer the VTID if present
  @VTID(21)
  void setUseLongFileNames(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "CheckIfOfficeIsHTMLEditor"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1905) //= 0x771. The runtime will prefer the VTID if present
  @VTID(22)
  boolean getCheckIfOfficeIsHTMLEditor();


  /**
   * <p>
   * Setter method for the COM property "CheckIfOfficeIsHTMLEditor"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1905) //= 0x771. The runtime will prefer the VTID if present
  @VTID(23)
  void setCheckIfOfficeIsHTMLEditor(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DownloadComponents"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1906) //= 0x772. The runtime will prefer the VTID if present
  @VTID(24)
  boolean getDownloadComponents();


  /**
   * <p>
   * Setter method for the COM property "DownloadComponents"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1906) //= 0x772. The runtime will prefer the VTID if present
  @VTID(25)
  void setDownloadComponents(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "RelyOnVML"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1907) //= 0x773. The runtime will prefer the VTID if present
  @VTID(26)
  boolean getRelyOnVML();


  /**
   * <p>
   * Setter method for the COM property "RelyOnVML"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1907) //= 0x773. The runtime will prefer the VTID if present
  @VTID(27)
  void setRelyOnVML(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "AllowPNG"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1908) //= 0x774. The runtime will prefer the VTID if present
  @VTID(28)
  boolean getAllowPNG();


  /**
   * <p>
   * Setter method for the COM property "AllowPNG"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1908) //= 0x774. The runtime will prefer the VTID if present
  @VTID(29)
  void setAllowPNG(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ScreenSize"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoScreenSize
   */

  @DISPID(1909) //= 0x775. The runtime will prefer the VTID if present
  @VTID(30)
  com.exceljava.com4j.office.MsoScreenSize getScreenSize();


  /**
   * <p>
   * Setter method for the COM property "ScreenSize"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoScreenSize parameter.
   */

  @DISPID(1909) //= 0x775. The runtime will prefer the VTID if present
  @VTID(31)
  void setScreenSize(
    com.exceljava.com4j.office.MsoScreenSize rhs);


  /**
   * <p>
   * Getter method for the COM property "PixelsPerInch"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1910) //= 0x776. The runtime will prefer the VTID if present
  @VTID(32)
  int getPixelsPerInch();


  /**
   * <p>
   * Setter method for the COM property "PixelsPerInch"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(1910) //= 0x776. The runtime will prefer the VTID if present
  @VTID(33)
  void setPixelsPerInch(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "LocationOfComponents"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1911) //= 0x777. The runtime will prefer the VTID if present
  @VTID(34)
  java.lang.String getLocationOfComponents();


  /**
   * <p>
   * Setter method for the COM property "LocationOfComponents"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1911) //= 0x777. The runtime will prefer the VTID if present
  @VTID(35)
  void setLocationOfComponents(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Encoding"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoEncoding
   */

  @DISPID(1822) //= 0x71e. The runtime will prefer the VTID if present
  @VTID(36)
  com.exceljava.com4j.office.MsoEncoding getEncoding();


  /**
   * <p>
   * Setter method for the COM property "Encoding"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoEncoding parameter.
   */

  @DISPID(1822) //= 0x71e. The runtime will prefer the VTID if present
  @VTID(37)
  void setEncoding(
    com.exceljava.com4j.office.MsoEncoding rhs);


  /**
   * <p>
   * Getter method for the COM property "AlwaysSaveInDefaultEncoding"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1912) //= 0x778. The runtime will prefer the VTID if present
  @VTID(38)
  boolean getAlwaysSaveInDefaultEncoding();


  /**
   * <p>
   * Setter method for the COM property "AlwaysSaveInDefaultEncoding"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1912) //= 0x778. The runtime will prefer the VTID if present
  @VTID(39)
  void setAlwaysSaveInDefaultEncoding(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Fonts"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.WebPageFonts
   */

  @DISPID(1913) //= 0x779. The runtime will prefer the VTID if present
  @VTID(40)
  com.exceljava.com4j.office.WebPageFonts getFonts();


  @VTID(40)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.WebPageFonts.class})
  com.exceljava.com4j.office.WebPageFont getFonts(
    com.exceljava.com4j.office.MsoCharacterSet index);

  /**
   * <p>
   * Getter method for the COM property "FolderSuffix"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1914) //= 0x77a. The runtime will prefer the VTID if present
  @VTID(41)
  java.lang.String getFolderSuffix();


  /**
   * <p>
   * Getter method for the COM property "TargetBrowser"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTargetBrowser
   */

  @DISPID(2179) //= 0x883. The runtime will prefer the VTID if present
  @VTID(42)
  com.exceljava.com4j.office.MsoTargetBrowser getTargetBrowser();


  /**
   * <p>
   * Setter method for the COM property "TargetBrowser"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoTargetBrowser parameter.
   */

  @DISPID(2179) //= 0x883. The runtime will prefer the VTID if present
  @VTID(43)
  void setTargetBrowser(
    com.exceljava.com4j.office.MsoTargetBrowser rhs);


  /**
   * <p>
   * Getter method for the COM property "SaveNewWebPagesAsWebArchives"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2180) //= 0x884. The runtime will prefer the VTID if present
  @VTID(44)
  boolean getSaveNewWebPagesAsWebArchives();


  /**
   * <p>
   * Setter method for the COM property "SaveNewWebPagesAsWebArchives"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2180) //= 0x884. The runtime will prefer the VTID if present
  @VTID(45)
  void setSaveNewWebPagesAsWebArchives(
    boolean rhs);


  // Properties:
}
