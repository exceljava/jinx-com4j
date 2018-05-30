package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000CD6A1-0000-0000-C000-000000000046}")
public interface SignatureSetup extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "ReadOnly"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(9)
  boolean getReadOnly();


  /**
   * <p>
   * Getter method for the COM property "Id"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(10)
  java.lang.String getId();


  /**
   * <p>
   * Getter method for the COM property "SignatureProvider"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(11)
  java.lang.String getSignatureProvider();


  /**
   * <p>
   * Getter method for the COM property "SuggestedSigner"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(12)
  java.lang.String getSuggestedSigner();


  /**
   * <p>
   * Setter method for the COM property "SuggestedSigner"
   * </p>
   * @param pbstr Mandatory java.lang.String parameter.
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(13)
  void setSuggestedSigner(
    java.lang.String pbstr);


  /**
   * <p>
   * Getter method for the COM property "SuggestedSignerLine2"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(14)
  java.lang.String getSuggestedSignerLine2();


  /**
   * <p>
   * Setter method for the COM property "SuggestedSignerLine2"
   * </p>
   * @param pbstr Mandatory java.lang.String parameter.
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(15)
  void setSuggestedSignerLine2(
    java.lang.String pbstr);


  /**
   * <p>
   * Getter method for the COM property "SuggestedSignerEmail"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(16)
  java.lang.String getSuggestedSignerEmail();


  /**
   * <p>
   * Setter method for the COM property "SuggestedSignerEmail"
   * </p>
   * @param pbstr Mandatory java.lang.String parameter.
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(17)
  void setSuggestedSignerEmail(
    java.lang.String pbstr);


  /**
   * <p>
   * Getter method for the COM property "SigningInstructions"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(18)
  java.lang.String getSigningInstructions();


  /**
   * <p>
   * Setter method for the COM property "SigningInstructions"
   * </p>
   * @param pbstr Mandatory java.lang.String parameter.
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(19)
  void setSigningInstructions(
    java.lang.String pbstr);


  /**
   * <p>
   * Getter method for the COM property "AllowComments"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(20)
  boolean getAllowComments();


  /**
   * <p>
   * Setter method for the COM property "AllowComments"
   * </p>
   * @param pvarf Mandatory boolean parameter.
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(21)
  void setAllowComments(
    boolean pvarf);


  /**
   * <p>
   * Getter method for the COM property "ShowSignDate"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(22)
  boolean getShowSignDate();


  /**
   * <p>
   * Setter method for the COM property "ShowSignDate"
   * </p>
   * @param pvarf Mandatory boolean parameter.
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(23)
  void setShowSignDate(
    boolean pvarf);


  /**
   * <p>
   * Getter method for the COM property "AdditionalXml"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(24)
  java.lang.String getAdditionalXml();


  /**
   * <p>
   * Setter method for the COM property "AdditionalXml"
   * </p>
   * @param pbstr Mandatory java.lang.String parameter.
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(25)
  void setAdditionalXml(
    java.lang.String pbstr);


  // Properties:
}
