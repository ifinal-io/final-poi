package org.ifinal.finalframework.poi;

import org.springframework.lang.NonNull;
import org.springframework.lang.Nullable;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.Date;
import java.util.HashSet;
import java.util.Objects;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;

/**
 * ObjectTypeHandler.
 *
 * @author likly
 * @version 1.0.0
 * @since 1.0.0
 */
public final class ObjectTypeHandler implements TypeHandler<Object> {

    private static final Set<Class<?>> NUMBERS = new HashSet<>(Arrays.asList(
        short.class, Short.class,
        int.class, Integer.class,
        long.class, Long.class,
        float.class, Float.class,
        double.class, Double.class
    ));

    private static final Set<Class<?>> BOOLEANS = new HashSet<>(Arrays.asList(boolean.class, Boolean.class));

    @Override
    public void setValue(@NonNull final Cell cell, @Nullable final Object value) {
        if (Objects.isNull(value)) {
            return;
        }

        Class<?> clazz = value.getClass();
        String text = value.toString();
        if (NUMBERS.contains(clazz)) {
            cell.setCellValue(Double.parseDouble(text));
        } else if (BOOLEANS.contains(clazz)) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else if (value instanceof LocalDateTime) {
            cell.setCellValue((LocalDateTime) value);
        } else if (value instanceof LocalDate) {
            cell.setCellValue((LocalDate) value);
        } else {

            if (text.startsWith("=")) {
                cell.setCellFormula(text.substring(1));
            } else {
                cell.setCellValue(text);
            }
        }
    }

}
