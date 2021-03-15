package org.ifinal.finalframework.poi;

/**
 * SpelExcelWriter.
 *
 * @author likly
 * @version 1.0.0
 * @since 1.0.0
 */
public class SpelExcelWriter extends AbstractExcelWriter {

    public SpelExcelWriter(final Excel excel) {
        super(excel, new SpelExcelGenerator());
    }

}
