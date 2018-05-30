package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0410-0000-0000-C000-000000000046}")
public interface SignatureSet extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(9)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610809345) //= 0x60030001. The runtime will prefer the VTID if present
  @VTID(10)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param iSig Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.Signature
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(11)
  @DefaultMethod
  com.exceljava.com4j.office.Signature getItem(
    int iSig);


  /**
   * @return  Returns a value of type com.exceljava.com4j.office.Signature
   */

  @DISPID(1610809347) //= 0x60030003. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.office.Signature add();


  /**
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  void commit();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610809349) //= 0x60030005. The runtime will prefer the VTID if present
  @VTID(14)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varSigProv is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addNonVisibleSignature(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Signature
   */

  @DISPID(1610809350) //= 0x60030006. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=1)
  com.exceljava.com4j.office.Signature addNonVisibleSignature();

  /**
   * @param varSigProv Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.Signature
   */

  @DISPID(1610809350) //= 0x60030006. The runtime will prefer the VTID if present
  @VTID(15)
  com.exceljava.com4j.office.Signature addNonVisibleSignature(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSigProv);


  /**
   * <p>
   * Getter method for the COM property "CanAddSignatureLine"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610809351) //= 0x60030007. The runtime will prefer the VTID if present
  @VTID(16)
  boolean getCanAddSignatureLine();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varSigProv is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addSignatureLine(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Signature
   */

  @DISPID(1610809352) //= 0x60030008. The runtime will prefer the VTID if present
  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=1)
  com.exceljava.com4j.office.Signature addSignatureLine();

  /**
   * @param varSigProv Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.Signature
   */

  @DISPID(1610809352) //= 0x60030008. The runtime will prefer the VTID if present
  @VTID(17)
  com.exceljava.com4j.office.Signature addSignatureLine(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSigProv);


  /**
   * <p>
   * Getter method for the COM property "Subset"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoSignatureSubset
   */

  @DISPID(1610809353) //= 0x60030009. The runtime will prefer the VTID if present
  @VTID(18)
  com.exceljava.com4j.office.MsoSignatureSubset getSubset();


  /**
   * <p>
   * Setter method for the COM property "Subset"
   * </p>
   * @param psubset Mandatory com.exceljava.com4j.office.MsoSignatureSubset parameter.
   */

  @DISPID(1610809353) //= 0x60030009. The runtime will prefer the VTID if present
  @VTID(19)
  void setSubset(
    com.exceljava.com4j.office.MsoSignatureSubset psubset);


  /**
   * <p>
   * Setter method for the COM property "ShowSignaturesPane"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1610809355) //= 0x6003000b. The runtime will prefer the VTID if present
  @VTID(20)
  void setShowSignaturesPane(
    boolean rhs);


  // Properties:
}
