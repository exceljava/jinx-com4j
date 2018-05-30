package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface PivotLine extends Com4jObject {
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
   * Getter method for the COM property "LineType"
   * </p>
   */

  @DISPID(2683)
  @PropGet
  com.exceljava.com4j.excel.XlPivotLineType getLineType();


  /**
   * <p>
   * Getter method for the COM property "Position"
   * </p>
   */

  @DISPID(133)
  @PropGet
  int getPosition();


  /**
   * <p>
   * Getter method for the COM property "PivotLineCells"
   * </p>
   */

  @DISPID(2684)
  @PropGet
  com.exceljava.com4j.excel.PivotLineCells getPivotLineCells();


  /**
   * <p>
   * Getter method for the COM property "PivotLineCellsFull"
   * </p>
   */

  @DISPID(3098)
  @PropGet
  com.exceljava.com4j.excel.PivotLineCells getPivotLineCellsFull();


  // Properties:
}
