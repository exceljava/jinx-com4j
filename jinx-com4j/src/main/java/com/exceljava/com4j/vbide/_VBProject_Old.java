package com.exceljava.com4j.vbide  ;

import com4j.*;

@IID("{0002E160-0000-0000-C000-000000000046}")
public interface _VBProject_Old extends com.exceljava.com4j.vbide._ProjectTemplate {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "HelpFile"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(116) //= 0x74. The runtime will prefer the VTID if present
  @VTID(9)
  java.lang.String getHelpFile();


  /**
   * <p>
   * Setter method for the COM property "HelpFile"
   * </p>
   * @param lpbstrHelpFile Mandatory java.lang.String parameter.
   */

  @DISPID(116) //= 0x74. The runtime will prefer the VTID if present
  @VTID(10)
  void setHelpFile(
    java.lang.String lpbstrHelpFile);


  /**
   * <p>
   * Getter method for the COM property "HelpContextID"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(117) //= 0x75. The runtime will prefer the VTID if present
  @VTID(11)
  int getHelpContextID();


  /**
   * <p>
   * Setter method for the COM property "HelpContextID"
   * </p>
   * @param lpdwContextID Mandatory int parameter.
   */

  @DISPID(117) //= 0x75. The runtime will prefer the VTID if present
  @VTID(12)
  void setHelpContextID(
    int lpdwContextID);


  /**
   * <p>
   * Getter method for the COM property "Description"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(118) //= 0x76. The runtime will prefer the VTID if present
  @VTID(13)
  java.lang.String getDescription();


  /**
   * <p>
   * Setter method for the COM property "Description"
   * </p>
   * @param lpbstrDescription Mandatory java.lang.String parameter.
   */

  @DISPID(118) //= 0x76. The runtime will prefer the VTID if present
  @VTID(14)
  void setDescription(
    java.lang.String lpbstrDescription);


  /**
   * <p>
   * Getter method for the COM property "Mode"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide.vbext_VBAMode
   */

  @DISPID(119) //= 0x77. The runtime will prefer the VTID if present
  @VTID(15)
  com.exceljava.com4j.vbide.vbext_VBAMode getMode();


  /**
   * <p>
   * Getter method for the COM property "References"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._References
   */

  @DISPID(120) //= 0x78. The runtime will prefer the VTID if present
  @VTID(16)
  com.exceljava.com4j.vbide._References getReferences();


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(121) //= 0x79. The runtime will prefer the VTID if present
  @VTID(17)
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param lpbstrName Mandatory java.lang.String parameter.
   */

  @DISPID(121) //= 0x79. The runtime will prefer the VTID if present
  @VTID(18)
  void setName(
    java.lang.String lpbstrName);


  /**
   * <p>
   * Getter method for the COM property "VBE"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide.VBE
   */

  @DISPID(122) //= 0x7a. The runtime will prefer the VTID if present
  @VTID(19)
  com.exceljava.com4j.vbide.VBE getVBE();


  /**
   * <p>
   * Getter method for the COM property "Collection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._VBProjects
   */

  @DISPID(123) //= 0x7b. The runtime will prefer the VTID if present
  @VTID(20)
  com.exceljava.com4j.vbide._VBProjects getCollection();


  /**
   * <p>
   * Getter method for the COM property "Protection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide.vbext_ProjectProtection
   */

  @DISPID(131) //= 0x83. The runtime will prefer the VTID if present
  @VTID(21)
  com.exceljava.com4j.vbide.vbext_ProjectProtection getProtection();


  /**
   * <p>
   * Getter method for the COM property "Saved"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(133) //= 0x85. The runtime will prefer the VTID if present
  @VTID(22)
  boolean getSaved();


  /**
   * <p>
   * Getter method for the COM property "VBComponents"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._VBComponents
   */

  @DISPID(135) //= 0x87. The runtime will prefer the VTID if present
  @VTID(23)
  com.exceljava.com4j.vbide._VBComponents getVBComponents();


  // Properties:
}
