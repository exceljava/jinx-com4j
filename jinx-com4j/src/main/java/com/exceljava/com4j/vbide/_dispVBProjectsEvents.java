package com.exceljava.com4j.vbide  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface _dispVBProjectsEvents extends Com4jObject {
  // Methods:
  /**
   * @param vbProject Mandatory com.exceljava.com4j.vbide._VBProject parameter.
   */

  @DISPID(1)
  void itemAdded(
    com.exceljava.com4j.vbide._VBProject vbProject);


  /**
   * @param vbProject Mandatory com.exceljava.com4j.vbide._VBProject parameter.
   */

  @DISPID(2)
  void itemRemoved(
    com.exceljava.com4j.vbide._VBProject vbProject);


  /**
   * @param vbProject Mandatory com.exceljava.com4j.vbide._VBProject parameter.
   * @param oldName Mandatory java.lang.String parameter.
   */

  @DISPID(3)
  void itemRenamed(
    com.exceljava.com4j.vbide._VBProject vbProject,
    java.lang.String oldName);


  /**
   * @param vbProject Mandatory com.exceljava.com4j.vbide._VBProject parameter.
   */

  @DISPID(4)
  void itemActivated(
    com.exceljava.com4j.vbide._VBProject vbProject);


  // Properties:
}
