package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{4CAC6328-B9B0-11D3-8D59-0050048384E3}")
public interface ILicWizExternal extends Com4jObject {
  // Methods:
  /**
   * @param punkHtmlDoc Mandatory com4j.Com4jObject parameter.
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(7)
  void printHtmlDocument(
    com4j.Com4jObject punkHtmlDoc);


  /**
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(8)
  void invokeDateTimeApplet();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter pFormat is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * formatDate(date, "");
   * </code>
   * </p>
   * @param date Mandatory java.util.Date parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(9)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.String.class}, nativeType = {NativeType.BSTR}, variantType = {Variant.Type.VT_BSTR}, literal = {""})
  @ReturnValue(index=2)
  java.lang.String formatDate(
    java.util.Date date);

  /**
   * @param date Mandatory java.util.Date parameter.
   * @param pFormat Optional parameter. Default value is ""
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(9)
  java.lang.String formatDate(
    java.util.Date date,
    @Optional @DefaultValue("") java.lang.String pFormat);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter pvarId is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * showHelp(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT_ByRef}, variantType = {Variant.Type.NO_TYPE}, literal = {"80020004"})
  void showHelp();

  /**
   * @param pvarId Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(10)
  void showHelp(
    @Optional java.lang.Object pvarId);


  /**
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(11)
  void terminate();


  /**
   * @param bpc Mandatory int parameter.
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(12)
  void disableVORWReminder(
    int bpc);


  /**
   * @param bstrReceipt Mandatory java.lang.String parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(13)
  java.lang.String saveReceipt(
    java.lang.String bstrReceipt);


  /**
   * @param bstrUrl Mandatory java.lang.String parameter.
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(14)
  void openInDefaultBrowser(
    java.lang.String bstrUrl);


  /**
   * @param bstrText Mandatory java.lang.String parameter.
   * @param bstrButtons Mandatory java.lang.String parameter.
   * @param bstrIcon Mandatory java.lang.String parameter.
   * @return  Returns a value of type int
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(15)
  int msoAlert(
    java.lang.String bstrText,
    java.lang.String bstrButtons,
    java.lang.String bstrIcon);


  /**
   * @param bstrKey Mandatory java.lang.String parameter.
   * @param fMORW Mandatory int parameter.
   * @return  Returns a value of type int
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(16)
  int depositPidKey(
    java.lang.String bstrKey,
    int fMORW);


  /**
   * @param bstrMessage Mandatory java.lang.String parameter.
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(17)
  void writeLog(
    java.lang.String bstrMessage);


  /**
   * @param bstrProductCode Mandatory java.lang.String parameter.
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(18)
  void resignDpc(
    java.lang.String bstrProductCode);


  /**
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(19)
  void resetPID();


  /**
   * @param dx Mandatory int parameter.
   * @param dy Mandatory int parameter.
   */

  @DISPID(14) //= 0xe. The runtime will prefer the VTID if present
  @VTID(20)
  void setDialogSize(
    int dx,
    int dy);


  /**
   * @param lMode Mandatory int parameter.
   * @return  Returns a value of type int
   */

  @DISPID(15) //= 0xf. The runtime will prefer the VTID if present
  @VTID(21)
  int verifyClock(
    int lMode);


  /**
   * @param pdispSelect Mandatory com4j.Com4jObject parameter.
   */

  @DISPID(16) //= 0x10. The runtime will prefer the VTID if present
  @VTID(22)
  void sortSelectOptions(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject pdispSelect);


  /**
   */

  @DISPID(17) //= 0x11. The runtime will prefer the VTID if present
  @VTID(23)
  void internetDisconnect();


  /**
   * @return  Returns a value of type int
   */

  @DISPID(18) //= 0x12. The runtime will prefer the VTID if present
  @VTID(24)
  int getConnectedState();


  /**
   * <p>
   * Getter method for the COM property "Context"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(20) //= 0x14. The runtime will prefer the VTID if present
  @VTID(25)
  int getContext();


  /**
   * <p>
   * Getter method for the COM property "Validator"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(21) //= 0x15. The runtime will prefer the VTID if present
  @VTID(26)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getValidator();


  /**
   * <p>
   * Getter method for the COM property "LicAgent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(22) //= 0x16. The runtime will prefer the VTID if present
  @VTID(27)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getLicAgent();


  /**
   * <p>
   * Getter method for the COM property "CountryInfo"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(23) //= 0x17. The runtime will prefer the VTID if present
  @VTID(28)
  java.lang.String getCountryInfo();


  /**
   * <p>
   * Setter method for the COM property "WizardVisible"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(24) //= 0x18. The runtime will prefer the VTID if present
  @VTID(29)
  void setWizardVisible(
    int rhs);


  /**
   * <p>
   * Setter method for the COM property "WizardTitle"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(25) //= 0x19. The runtime will prefer the VTID if present
  @VTID(30)
  void setWizardTitle(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "AnimationEnabled"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(26) //= 0x1a. The runtime will prefer the VTID if present
  @VTID(31)
  int getAnimationEnabled();


  /**
   * <p>
   * Setter method for the COM property "CurrentHelpId"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(27) //= 0x1b. The runtime will prefer the VTID if present
  @VTID(32)
  void setCurrentHelpId(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "OfficeOnTheWebUrl"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(28) //= 0x1c. The runtime will prefer the VTID if present
  @VTID(33)
  java.lang.String getOfficeOnTheWebUrl();


  // Properties:
}
