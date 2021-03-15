package org.ifinal.finalframework.poi;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

/**
 * ColorsTest.
 *
 * @author likly
 * @version 1.0.0
 * @since 1.0.0
 */
class ColorsTest {

    @Test
    void decode() {

        assertEquals(255, Colors.decode("RED").getRed());

    }

}
