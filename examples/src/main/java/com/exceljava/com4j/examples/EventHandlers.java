package com.exceljava.com4j.examples;

import com.exceljava.com4j.excel._Chart;
import com.exceljava.com4j.excel._Worksheet;
import com.exceljava.jinx.events.OnSheetActivate;
import com4j.COM4J;
import com4j.Com4jObject;
import com4j.util.ComObjectCollector;

import java.util.logging.Logger;

/**
 * Example Excel event handlers that use com4j to access the
 * COM objects passed to them by Excel.
 *
 * Excel event annotations require Jinx 3.0 or later.
 */
public class EventHandlers {
    private static final Logger log = Logger.getLogger(EventHandlers.class.getName());

    /**
     * Called by Excel whenever the active sheet changes.
     *
     * The sheet is converted from an IUnknown to a Com4jObject by jinx-com4j.
     * It is disposed of automatically once this method returns, so it
     * should not be kept after the call.
     *
     * @param sheet The sheet that has been activated. This may be a
     *              worksheet or a chart sheet.
     */
    @OnSheetActivate
    public static void onSheetActivate(Com4jObject sheet) {
        // Collect the COM objects created by this method so that they can
        // all be disposed of before returning
        ComObjectCollector collector = new ComObjectCollector();
        COM4J.addListener(collector);

        try {
            String name;
            if (sheet.is(_Worksheet.class)) {
                name = sheet.queryInterface(_Worksheet.class).getName();
            } else if (sheet.is(_Chart.class)) {
                name = sheet.queryInterface(_Chart.class).getName();
            } else {
                name = "<unknown>";
            }

            log.info("jinx-com4j EventHandlers example: Active sheet changed to '" + name + "'");
        }
        finally {
            collector.disposeAll();
            COM4J.removeListener(collector);
        }
    }
}