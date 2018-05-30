package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0411-0000-0000-C000-000000000046}")
public interface Signature extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Signer"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610809344) //= 0x60030000. The runtime will prefer the VTID if present
  @VTID(9)
  java.lang.String getSigner();


  /**
   * <p>
   * Getter method for the COM property "Issuer"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610809345) //= 0x60030001. The runtime will prefer the VTID if present
  @VTID(10)
  java.lang.String getIssuer();


  /**
   * <p>
   * Getter method for the COM property "ExpireDate"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1610809346) //= 0x60030002. The runtime will prefer the VTID if present
  @VTID(11)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getExpireDate();


  /**
   * <p>
   * Getter method for the COM property "IsValid"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610809347) //= 0x60030003. The runtime will prefer the VTID if present
  @VTID(12)
  boolean getIsValid();


  /**
   * <p>
   * Getter method for the COM property "AttachCertificate"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  boolean getAttachCertificate();


  /**
   * <p>
   * Setter method for the COM property "AttachCertificate"
   * </p>
   * @param pvarfAttach Mandatory boolean parameter.
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(14)
  void setAttachCertificate(
    boolean pvarfAttach);


  /**
   */

  @DISPID(1610809350) //= 0x60030006. The runtime will prefer the VTID if present
  @VTID(15)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610809351) //= 0x60030007. The runtime will prefer the VTID if present
  @VTID(16)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "IsCertificateExpired"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610809352) //= 0x60030008. The runtime will prefer the VTID if present
  @VTID(17)
  boolean getIsCertificateExpired();


  /**
   * <p>
   * Getter method for the COM property "IsCertificateRevoked"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610809353) //= 0x60030009. The runtime will prefer the VTID if present
  @VTID(18)
  boolean getIsCertificateRevoked();


  /**
   * <p>
   * Getter method for the COM property "SignDate"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1610809354) //= 0x6003000a. The runtime will prefer the VTID if present
  @VTID(19)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getSignDate();


  /**
   * <p>
   * Getter method for the COM property "IsSigned"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610809355) //= 0x6003000b. The runtime will prefer the VTID if present
  @VTID(20)
  boolean getIsSigned();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varSigImg is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varDelSuggSigner is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varDelSuggSignerLine2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varDelSuggSignerEmail is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sign(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1610809356) //= 0x6003000c. The runtime will prefer the VTID if present
  @VTID(21)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void sign();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varDelSuggSigner is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varDelSuggSignerLine2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varDelSuggSignerEmail is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sign(varSigImg, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param varSigImg Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610809356) //= 0x6003000c. The runtime will prefer the VTID if present
  @VTID(21)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void sign(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSigImg);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varDelSuggSignerLine2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varDelSuggSignerEmail is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sign(varSigImg, varDelSuggSigner, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param varSigImg Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varDelSuggSigner Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610809356) //= 0x6003000c. The runtime will prefer the VTID if present
  @VTID(21)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void sign(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSigImg,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varDelSuggSigner);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varDelSuggSignerEmail is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * sign(varSigImg, varDelSuggSigner, varDelSuggSignerLine2, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param varSigImg Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varDelSuggSigner Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varDelSuggSignerLine2 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610809356) //= 0x6003000c. The runtime will prefer the VTID if present
  @VTID(21)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void sign(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSigImg,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varDelSuggSigner,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varDelSuggSignerLine2);

  /**
   * @param varSigImg Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varDelSuggSigner Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varDelSuggSignerLine2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varDelSuggSignerEmail Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610809356) //= 0x6003000c. The runtime will prefer the VTID if present
  @VTID(21)
  void sign(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSigImg,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varDelSuggSigner,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varDelSuggSignerLine2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varDelSuggSignerEmail);


  /**
   * <p>
   * Getter method for the COM property "Details"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SignatureInfo
   */

  @DISPID(1610809357) //= 0x6003000d. The runtime will prefer the VTID if present
  @VTID(22)
  com.exceljava.com4j.office.SignatureInfo getDetails();


  /**
   */

  @DISPID(1610809358) //= 0x6003000e. The runtime will prefer the VTID if present
  @VTID(23)
  void showDetails();


  /**
   * <p>
   * Getter method for the COM property "CanSetup"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610809359) //= 0x6003000f. The runtime will prefer the VTID if present
  @VTID(24)
  boolean getCanSetup();


  /**
   * <p>
   * Getter method for the COM property "Setup"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SignatureSetup
   */

  @DISPID(1610809360) //= 0x60030010. The runtime will prefer the VTID if present
  @VTID(25)
  com.exceljava.com4j.office.SignatureSetup getSetup();


  /**
   * <p>
   * Getter method for the COM property "IsSignatureLine"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610809361) //= 0x60030011. The runtime will prefer the VTID if present
  @VTID(26)
  boolean getIsSignatureLine();


  /**
   * <p>
   * Getter method for the COM property "SignatureLineShape"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610809362) //= 0x60030012. The runtime will prefer the VTID if present
  @VTID(27)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getSignatureLineShape();


  /**
   * <p>
   * Getter method for the COM property "SortHint"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610809363) //= 0x60030013. The runtime will prefer the VTID if present
  @VTID(28)
  int getSortHint();


  // Properties:
}
