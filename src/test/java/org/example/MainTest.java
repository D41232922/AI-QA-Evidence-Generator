package org.example;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class MainTest {
    @Test
    void detectsScreenshotWord() {
        assertTrue(Main.containsScreenshot("Please take a screenshot here"));
        assertTrue(Main.containsScreenshot("SCREENSHOT"));
        assertFalse(Main.containsScreenshot("Please take a picture"));
    }

    @Test
    void formatsSecondsAsTimestamp() {
        assertEquals("00:00:12.345", Main.formatSeconds(12.345));
        assertEquals("01:02:03.004", Main.formatSeconds(3723.004));
    }
}
