package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C03C4-0000-0000-C000-000000000046}")
public interface IBlogExtensibility extends Com4jObject {
  // Methods:
  /**
   * @param blogProvider Mandatory Holder<java.lang.String> parameter.
   * @param friendlyName Mandatory Holder<java.lang.String> parameter.
   * @param categorySupport Mandatory Holder<com.exceljava.com4j.office.MsoBlogCategorySupport> parameter.
   * @param padding Mandatory Holder<Boolean> parameter.
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(7)
  void blogProviderProperties(
    Holder<java.lang.String> blogProvider,
    Holder<java.lang.String> friendlyName,
    Holder<com.exceljava.com4j.office.MsoBlogCategorySupport> categorySupport,
    Holder<Boolean> padding);


  /**
   * @param account Mandatory java.lang.String parameter.
   * @param parentWindow Mandatory int parameter.
   * @param document Mandatory com4j.Com4jObject parameter.
   * @param newAccount Mandatory boolean parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(8)
  boolean setupBlogAccount(
    java.lang.String account,
    int parentWindow,
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject document,
    boolean newAccount);


  /**
   * @param account Mandatory java.lang.String parameter.
   * @param parentWindow Mandatory int parameter.
   * @param document Mandatory com4j.Com4jObject parameter.
   * @param blogNames Mandatory Holder<java.lang.String[]> parameter.
   * @param blogIDs Mandatory Holder<java.lang.String[]> parameter.
   * @param blogURLs Mandatory Holder<java.lang.String[]> parameter.
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(9)
  void getUserBlogs(
    java.lang.String account,
    int parentWindow,
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject document,
    Holder<java.lang.String[]> blogNames,
    Holder<java.lang.String[]> blogIDs,
    Holder<java.lang.String[]> blogURLs);


  /**
   * @param account Mandatory java.lang.String parameter.
   * @param parentWindow Mandatory int parameter.
   * @param document Mandatory com4j.Com4jObject parameter.
   * @param postTitles Mandatory Holder<java.lang.String[]> parameter.
   * @param postDates Mandatory Holder<java.lang.String[]> parameter.
   * @param postIDs Mandatory Holder<java.lang.String[]> parameter.
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(10)
  void getRecentPosts(
    java.lang.String account,
    int parentWindow,
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject document,
    Holder<java.lang.String[]> postTitles,
    Holder<java.lang.String[]> postDates,
    Holder<java.lang.String[]> postIDs);


  /**
   * @param account Mandatory java.lang.String parameter.
   * @param postID Mandatory java.lang.String parameter.
   * @param parentWindow Mandatory int parameter.
   * @param xHTML Mandatory Holder<java.lang.String> parameter.
   * @param title Mandatory Holder<java.lang.String> parameter.
   * @param datePosted Mandatory Holder<java.lang.String> parameter.
   * @param categories Mandatory Holder<java.lang.String[]> parameter.
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(11)
  void open(
    java.lang.String account,
    java.lang.String postID,
    int parentWindow,
    Holder<java.lang.String> xHTML,
    Holder<java.lang.String> title,
    Holder<java.lang.String> datePosted,
    Holder<java.lang.String[]> categories);


  /**
   * @param account Mandatory java.lang.String parameter.
   * @param parentWindow Mandatory int parameter.
   * @param document Mandatory com4j.Com4jObject parameter.
   * @param xHTML Mandatory java.lang.String parameter.
   * @param title Mandatory java.lang.String parameter.
   * @param dateTime Mandatory java.lang.String parameter.
   * @param categories Mandatory java.lang.String[] parameter.
   * @param draft Mandatory boolean parameter.
   * @param postID Mandatory Holder<java.lang.String> parameter.
   * @param publishMessage Mandatory Holder<java.lang.String> parameter.
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(12)
  void publishPost(
    java.lang.String account,
    int parentWindow,
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject document,
    java.lang.String xHTML,
    java.lang.String title,
    java.lang.String dateTime,
    java.lang.String[] categories,
    boolean draft,
    Holder<java.lang.String> postID,
    Holder<java.lang.String> publishMessage);


  /**
   * @param account Mandatory java.lang.String parameter.
   * @param parentWindow Mandatory int parameter.
   * @param document Mandatory com4j.Com4jObject parameter.
   * @param postID Mandatory java.lang.String parameter.
   * @param xHTML Mandatory java.lang.String parameter.
   * @param title Mandatory java.lang.String parameter.
   * @param dateTime Mandatory java.lang.String parameter.
   * @param categories Mandatory java.lang.String[] parameter.
   * @param draft Mandatory boolean parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(13)
  java.lang.String republishPost(
    java.lang.String account,
    int parentWindow,
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject document,
    java.lang.String postID,
    java.lang.String xHTML,
    java.lang.String title,
    java.lang.String dateTime,
    java.lang.String[] categories,
    boolean draft);


  /**
   * @param account Mandatory java.lang.String parameter.
   * @param parentWindow Mandatory int parameter.
   * @param document Mandatory com4j.Com4jObject parameter.
   * @return  Returns a value of type java.lang.String[]
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(14)
  java.lang.String[] getCategories(
    java.lang.String account,
    int parentWindow,
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject document);


  // Properties:
}
