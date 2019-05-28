package com.exceljava.com4j.examples;

import com.exceljava.com4j.JinxBridge;
import com.exceljava.com4j.excel.*;
import com.exceljava.jinx.*;
import com4j.Com4jObject;

import java.util.Random;

/**
 * Example worksheet functions that call back into Excel using com4j.
 */
public class WorksheetFunctions {
    private final ExcelAddIn xl;

    public WorksheetFunctions(ExcelAddIn xl) {
        this.xl = xl;
    }

    @ExcelFunction(
            value="com4j.coloredTableExample",
            autoResize = true,
            isMacroType = true
    )
    public Object[][] coloredTableExample(int numRows, int numCols) {
        // Make sure the table is not zero sized
        int usedNumRows = Math.max(numRows, 0);
        int usedNumCols = Math.max(numCols, 1);

        // Create the array for the result
        Object[][] result = new Object[usedNumRows + 1][usedNumCols];

        // Set the header
        for (int c=0; c < usedNumCols; c++) {
            char name = (char)((c % 26) + 'A');
            result[0][c] = String.valueOf(name);
        }

        // Fill the rest of the table with random data
        Random rand = new Random();
        for (int r=1; r <= usedNumRows; r++) {
            for (int c=0; c < usedNumCols; c++) {
                result[r][c] = rand.nextDouble();
            }
        }

        // Get the address of the calling range
        ExcelReference caller = xl.getCaller();
        String address = caller.getAddress();

        // Before returning, schedule a call to color the returned table
        xl.schedule(() -> {
            _Application app = JinxBridge.getApplication(xl);
            Range range = app.getRange(address);

            // Excel ranges are indexed from 1
            Range topLeft = ((Com4jObject)range.getItem(1, 1)).queryInterface(Range.class);
            Range topRight = ((Com4jObject)range.getItem(1, usedNumCols)).queryInterface(Range.class);

            // Set the background and font color of the header
            Range header = app.getRange(topLeft, topRight);
            header.getInterior().setColor(0xC47244);
            header.getFont().setColor(0xFFFFFF);

            // Set the background and font color of the rows
            for (int i=2; i<=usedNumRows; i++) {
                Range left = ((Com4jObject)range.getItem(i, 1)).queryInterface(Range.class);
                Range right = ((Com4jObject)range.getItem(i, usedNumCols)).queryInterface(Range.class);
                Range row = app.getRange(left, right);
                row.getInterior().setColor(i % 2 == 0 ? 0xFAF1EA : 0xFFFFFF);
                row.getFont().setColor(0x000000);
            }
        });

        return result;
    }
}
