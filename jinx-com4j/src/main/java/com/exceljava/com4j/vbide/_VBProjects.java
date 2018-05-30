package com.exceljava.com4j.vbide  ;

import com4j.*;

@IID("{EEE00919-E393-11D1-BB03-00C04FB6C4A6}")
public interface _VBProjects extends com.exceljava.com4j.vbide._VBProjects_Old {
  // Methods:
  /**
   * @param type Mandatory com.exceljava.com4j.vbide.vbext_ProjectType parameter.
   * @return  Returns a value of type com.exceljava.com4j.vbide._VBProject
   */

  @DISPID(137) //= 0x89. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.vbide._VBProject add(
    com.exceljava.com4j.vbide.vbext_ProjectType type);


  /**
   * @param lpc Mandatory com.exceljava.com4j.vbide._VBProject parameter.
   */

  @DISPID(138) //= 0x8a. The runtime will prefer the VTID if present
  @VTID(13)
  void remove(
    com.exceljava.com4j.vbide._VBProject lpc);


  /**
   * @param bstrPath Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.vbide._VBProject
   */

  @DISPID(139) //= 0x8b. The runtime will prefer the VTID if present
  @VTID(14)
  com.exceljava.com4j.vbide._VBProject open(
    java.lang.String bstrPath);


  // Properties:
}
