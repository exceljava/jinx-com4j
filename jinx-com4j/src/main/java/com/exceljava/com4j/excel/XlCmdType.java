package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlCmdType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlCmdCube(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlCmdSql(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlCmdTable(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlCmdDefault(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlCmdList(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlCmdTableCollection(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlCmdExcel(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlCmdDAX(8),
  ;

  private final int value;
  XlCmdType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
