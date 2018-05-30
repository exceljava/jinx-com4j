package com.exceljava.com4j.vbide  ;

import com4j.*;

@IID("{EEE00915-E393-11D1-BB03-00C04FB6C4A6}")
public interface _VBProject extends com.exceljava.com4j.vbide._VBProject_Old {
  // Methods:
  /**
   * @param fileName Mandatory java.lang.String parameter.
   */

  @DISPID(138) //= 0x8a. The runtime will prefer the VTID if present
  @VTID(24)
  void saveAs(
    java.lang.String fileName);


  /**
   */

  @DISPID(139) //= 0x8b. The runtime will prefer the VTID if present
  @VTID(25)
  void makeCompiledFile();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide.vbext_ProjectType
   */

  @DISPID(140) //= 0x8c. The runtime will prefer the VTID if present
  @VTID(26)
  com.exceljava.com4j.vbide.vbext_ProjectType getType();


  /**
   * <p>
   * Getter method for the COM property "FileName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(141) //= 0x8d. The runtime will prefer the VTID if present
  @VTID(27)
  java.lang.String getFileName();


  /**
   * <p>
   * Getter method for the COM property "BuildFileName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(142) //= 0x8e. The runtime will prefer the VTID if present
  @VTID(28)
  java.lang.String getBuildFileName();


  /**
   * <p>
   * Setter method for the COM property "BuildFileName"
   * </p>
   * @param lpbstrBldFName Mandatory java.lang.String parameter.
   */

  @DISPID(142) //= 0x8e. The runtime will prefer the VTID if present
  @VTID(29)
  void setBuildFileName(
    java.lang.String lpbstrBldFName);


  // Properties:
}
