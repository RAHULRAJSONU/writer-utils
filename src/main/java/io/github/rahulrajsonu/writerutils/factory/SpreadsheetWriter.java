package io.github.rahulrajsonu.writerutils.factory;

import io.github.rahulrajsonu.writerutils.annotation.XlsxCompositeField;
import io.github.rahulrajsonu.writerutils.annotation.XlsxSheet;
import io.github.rahulrajsonu.writerutils.annotation.XlsxSingleField;
import io.github.rahulrajsonu.writerutils.service.Writer;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayOutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.sql.Date;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.List;

public class SpreadsheetWriter implements Writer {

    private static final Logger logger = LoggerFactory.getLogger(SpreadsheetWriter.class);
    @Override
    public <T> byte[] write(List<T> data, String sheetName) {

        long start = System.currentTimeMillis();

        Class<?> clazz = data.get(0).getClass();

        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream();){
            // setting up the basic styles for the workbook
            Font boldFont = getBoldFont(workbook);
            Font genericFont = getGenericFont(workbook);
            CellStyle headerStyle = getLeftAlignedCellStyle(workbook, boldFont);
            CellStyle currencyStyle = setCurrencyCellStyle(workbook);
            CellStyle centerAlignedStyle = getCenterAlignedCellStyle(workbook);
            CellStyle genericStyle = getLeftAlignedCellStyle(workbook, genericFont);
            Sheet sheet = workbook.createSheet(sheetName);

            // get the metadata for each field of the POJO class into a list
            List<XlsxField> xlsColumnFields = getFieldNamesForClass(clazz);
            String[] columnTitles = xlsColumnFields.stream().map(XlsxField::getColumnHeader).toArray(String[]::new);
            int tempRowNo = 0;
            int recordBeginRowNo = 0;
            int recordEndRowNo = 0;

            // set spreadsheet titles
            Row mainRow = sheet.createRow(tempRowNo);
            Cell columnTitleCell;

            for (int i = 0; i < columnTitles.length; i++) {
                columnTitleCell = mainRow.createCell(i);
                columnTitleCell.setCellStyle(headerStyle);
                columnTitleCell.setCellValue(columnTitles[i]);

            }

            recordEndRowNo++;

//            looping the past dataset
            for (T record : data) {

                tempRowNo = recordEndRowNo;
                recordBeginRowNo = tempRowNo;
                mainRow = sheet.createRow(tempRowNo++);

                boolean isFirstValue;
                boolean isFirstRow;
                boolean isRowNoToDecrease = false;
                Method xlsMethod;
                Object xlsObjValue;
                List<Object> objValueList;

//                get max size of the record if its multiple row
                int maxListSize = getMaxListSize(record, xlsColumnFields, clazz);


//                looping through the fields of the current record
                for (XlsxField xlsColumnField : xlsColumnFields) {

//                    writing a single field
                    if (!xlsColumnField.isAnArray() && !xlsColumnField.isComposite()) {

                        writeSingleFieldRow(mainRow, xlsColumnField, clazz, currencyStyle, centerAlignedStyle, genericStyle,
                                record, workbook);

//                        overlooking the next field and adjusting the starting row
                        if (isNextColumnAnArray(xlsColumnFields, xlsColumnField, clazz, record)) {
                            isRowNoToDecrease = true;
                            tempRowNo = recordBeginRowNo + 1;
                        }

//                        writing an single array field
                    } else if (xlsColumnField.isAnArray() && !xlsColumnField.isComposite()) {

                        xlsMethod = getMethod(clazz, xlsColumnField);
                        xlsObjValue = xlsMethod.invoke(record, (Object[]) null);
                        objValueList = (List<Object>) xlsObjValue;
                        isFirstValue = true;

//                        looping through the items of the single array
                        for (Object objectValue : objValueList) {

                            Row childRow;

                            if (isFirstValue) {
                                childRow = mainRow;
                                writeArrayFieldRow(childRow, xlsColumnField, objectValue, currencyStyle, centerAlignedStyle,
                                        genericStyle, workbook);
                                isFirstValue = false;

                            } else if (isRowNoToDecrease) {
                                childRow = getOrCreateNextRow(sheet, tempRowNo++);
                                writeArrayFieldRow(childRow, xlsColumnField, objectValue, currencyStyle, centerAlignedStyle,
                                        genericStyle, workbook);
                                isRowNoToDecrease = false;

                            } else {
                                childRow = getOrCreateNextRow(sheet, tempRowNo++);
                                writeArrayFieldRow(childRow, xlsColumnField, objectValue, currencyStyle, centerAlignedStyle,
                                        genericStyle, workbook);
                            }
                        }

                        //                        overlooking the next field and adjusting the starting row
                        if (isNextColumnAnArray(xlsColumnFields, xlsColumnField, clazz, record)) {
                            isRowNoToDecrease = true;
                            tempRowNo = recordBeginRowNo + 1;
                        }

//                        writing a composite array field
                    } else if (xlsColumnField.isAnArray() && xlsColumnField.isComposite()) {

                        xlsMethod = getMethod(clazz, xlsColumnField);
                        xlsObjValue = xlsMethod.invoke(record, (Object[]) null);
                        objValueList = (List<Object>) xlsObjValue;
                        isFirstRow = true;

//                        looping through the items of the composite array
                        for (Object objectValue : objValueList) {

                            Row childRow;
                            List<XlsxField> xlsCompositeColumnFields = getFieldNamesForClass(objectValue.getClass());

                            if (isFirstRow) {
                                childRow = mainRow;
                                for (XlsxField xlsCompositeColumnField : xlsCompositeColumnFields) {
                                    writeCompositeFieldRow(objectValue, xlsCompositeColumnField, childRow, currencyStyle,
                                            centerAlignedStyle, genericStyle, workbook);
                                }
                                isFirstRow = false;

                            } else if (isRowNoToDecrease) {

                                childRow = getOrCreateNextRow(sheet, tempRowNo++);
                                for (XlsxField xlsCompositeColumnField : xlsCompositeColumnFields) {
                                    writeCompositeFieldRow(objectValue, xlsCompositeColumnField, childRow, currencyStyle,
                                            centerAlignedStyle, genericStyle, workbook);
                                }
                                isRowNoToDecrease = false;

                            } else {
                                childRow = getOrCreateNextRow(sheet, tempRowNo++);
                                for (XlsxField xlsCompositeColumnField : xlsCompositeColumnFields) {
                                    writeCompositeFieldRow(objectValue, xlsCompositeColumnField, childRow, currencyStyle,
                                            centerAlignedStyle, genericStyle, workbook);
                                }
                            }
                        }

//                        overlooking the next field and adjusting the starting row
                        if (isNextColumnAnArray(xlsColumnFields, xlsColumnField, clazz, record)) {
                            isRowNoToDecrease = true;
                            tempRowNo = recordBeginRowNo + 1;
                        }
                    }
                }

//                adjusting the ending row number for the current record
                recordEndRowNo = maxListSize + recordBeginRowNo;
            }

//            auto sizing the columns of the whole sheet
            autoSizeColumns(sheet, xlsColumnFields.size());

            workbook.write(out);
            logger.info("Xls file generated in [{}] seconds", processTime(start));
            return out.toByteArray();
        } catch (Exception e) {
            logger.info("Xls file write failed", e);
            throw new RuntimeException("Xls file write failed, error: "+e.getMessage());
        }
    }

    private void writeCompositeFieldRow(Object objectValue, XlsxField xlsCompositeColumnField, Row childRow,
                                        CellStyle currencyStyle, CellStyle centerAlignedStyle, CellStyle genericStyle,
                                        Workbook workbook)
            throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {

        Method nestedCompositeXlsMethod = getMethod(objectValue.getClass(), xlsCompositeColumnField);
        Object nestedCompositeValue = nestedCompositeXlsMethod.invoke(objectValue, (Object[]) null);
        Cell compositeNewCell = childRow.createCell(xlsCompositeColumnField.getCellIndex());
        setCellValue(compositeNewCell, nestedCompositeValue, currencyStyle, centerAlignedStyle, genericStyle, workbook);

    }

    private void writeArrayFieldRow(Row childRow, XlsxField xlsColumnField, Object objectValue,
                                    CellStyle currencyStyle, CellStyle centerAlignedStyle, CellStyle genericStyle, Workbook workbook) {
        Cell newCell = childRow.createCell(xlsColumnField.getCellIndex());
        setCellValue(newCell, objectValue, currencyStyle, centerAlignedStyle, genericStyle, workbook);
    }

    private <T> void writeSingleFieldRow(Row mainRow, XlsxField xlsColumnField, Class<?> clazz, CellStyle currencyStyle,
                                         CellStyle centerAlignedStyle, CellStyle genericStyle, T record, Workbook workbook)
            throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {

        Cell newCell = mainRow.createCell(xlsColumnField.getCellIndex());
        Method xlsMethod = getMethod(clazz, xlsColumnField);
        Object xlsObjValue = xlsMethod.invoke(record, (Object[]) null);
        setCellValue(newCell, xlsObjValue, currencyStyle, centerAlignedStyle, genericStyle, workbook);

    }

    private <T> boolean isNextColumnAnArray(List<XlsxField> xlsColumnFields, XlsxField xlsColumnField,
                                            Class<?> clazz, T record)
            throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {

        XlsxField nextXlsColumnField;
        int fieldsSize = xlsColumnFields.size();
        Method nestedXlsMethod;
        Object nestedObjValue;
        List<Object> nestedObjValueList;

        if (xlsColumnFields.indexOf(xlsColumnField) < (fieldsSize - 1)) {
            nextXlsColumnField = xlsColumnFields.get(xlsColumnFields.indexOf(xlsColumnField) + 1);
            if (nextXlsColumnField.isAnArray()) {
                nestedXlsMethod = getMethod(clazz, nextXlsColumnField);
                nestedObjValue = nestedXlsMethod.invoke(record, (Object[]) null);
                nestedObjValueList = (List<Object>) nestedObjValue;
                return nestedObjValueList.size() > 1;
            }
        }

        return xlsColumnFields.indexOf(xlsColumnField) == (fieldsSize - 1);

    }


    private void setCellValue(Cell cell, Object objValue, CellStyle currencyStyle, CellStyle centerAlignedStyle,
                              CellStyle genericStyle, Workbook workbook) {

        Hyperlink link = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
        CreationHelper createHelper = workbook.getCreationHelper();
        CellStyle dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat(
                createHelper.createDataFormat().getFormat("yyyy/MM/dd"));

        if (objValue != null) {
            if (objValue instanceof String) {
                String cellValue = (String) objValue;
                cell.setCellStyle(genericStyle);
                if (cellValue.contains("https://") || cellValue.contains("http://")) {
                    link.setAddress(cellValue);
                    cell.setCellValue(cellValue);
                    cell.setHyperlink(link);
                } else {
                    cell.setCellValue(cellValue);
                }
            } else if (objValue instanceof Long) {
                cell.setCellValue((Long) objValue);
            } else if (objValue instanceof Integer) {
                cell.setCellValue((Integer) objValue);
            } else if (objValue instanceof Double) {
                Double cellValue = (Double) objValue;
                cell.setCellStyle(currencyStyle);
                cell.setCellValue(cellValue);
            } else if (objValue instanceof Boolean) {
                cell.setCellStyle(centerAlignedStyle);
                if (objValue.equals(true)) {
                    cell.setCellValue(1);
                } else {
                    cell.setCellValue(0);
                }
            } else if(objValue instanceof Date){
                Date date = (Date)objValue;
                cell.setCellValue(date);
                cell.setCellStyle(dateCellStyle);
            }
        }
    }

    private static List<XlsxField> getFieldNamesForClass(Class<?> clazz) {
        List<XlsxField> xlsColumnFields = new ArrayList<>();
        Field[] fields = clazz.getDeclaredFields();
        XlsxSheet xlsxSheet = clazz.getAnnotation(XlsxSheet.class);
        if(null != xlsxSheet) {
            for (Field field : fields) {

                XlsxField xlsColumnField = new XlsxField();
                if (Collection.class.isAssignableFrom(field.getType())) {
                    xlsColumnField.setAnArray(true);
                    XlsxCompositeField xlsCompositeField = field.getAnnotation(XlsxCompositeField.class);
                    if (xlsCompositeField != null) {
                        xlsColumnField.setCellIndexFrom(xlsCompositeField.from());
                        xlsColumnField.setCellIndexTo(xlsCompositeField.to());
                        xlsColumnField.setComposite(true);
                    } else {
                        XlsxSingleField xlsField = field.getAnnotation(XlsxSingleField.class);
                        xlsColumnField.setCellIndex(xlsField.columnIndex());
                        xlsColumnField.setColumnHeader(xlsField.columnHeader());
                    }
                } else {
                    XlsxSingleField xlsField = field.getAnnotation(XlsxSingleField.class);
                    xlsColumnField.setAnArray(false);
                    if (xlsField != null) {
                        xlsColumnField.setCellIndex(xlsField.columnIndex());
                        xlsColumnField.setColumnHeader(xlsField.columnHeader());
                        xlsColumnField.setComposite(false);
                    }
                }
                xlsColumnField.setFieldName(field.getName());
                xlsColumnFields.add(xlsColumnField);
            }
        }else {
            int columnIndex = 0;
            for (Field field : fields) {
                XlsxField xlsColumnField = new XlsxField();
                xlsColumnField.setAnArray(false);
                xlsColumnField.setCellIndex(columnIndex);
                xlsColumnField.setColumnHeader(field.getName());
                xlsColumnField.setComposite(false);
                xlsColumnField.setFieldName(field.getName());
                xlsColumnFields.add(xlsColumnField);
                columnIndex++;
            }
        }
        return xlsColumnFields;
    }

    private static String capitalize(String s) {
        if (s.length() == 0)
            return s;
        return s.substring(0, 1).toUpperCase() + s.substring(1);
    }


    private <T> int getMaxListSize(T record, List<XlsxField> xlsColumnFields, Class<? extends Object> aClass)
            throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {

        List<Integer> listSizes = new ArrayList<>();

        for (XlsxField xlsColumnField : xlsColumnFields) {
            if (xlsColumnField.isAnArray()) {
                Method method = getMethod(aClass, xlsColumnField);
                Object value = method.invoke(record, (Object[]) null);
                List<Object> objects = (List<Object>) value;
                if (objects.size() > 1) {
                    listSizes.add(objects.size());
                }
            }
        }

        if (listSizes.isEmpty()) {
            return 1;
        } else {
            return Collections.max(listSizes);
        }

    }

    private Method getMethod(Class<?> clazz, XlsxField xlsColumnField) throws NoSuchMethodException {
        Method method;
        try {
            method = clazz.getMethod("get" + capitalize(xlsColumnField.getFieldName()));
        } catch (NoSuchMethodException nme) {
            method = clazz.getMethod(xlsColumnField.getFieldName());
        }

        return method;
    }

    private long processTime(long start) {
        return (System.currentTimeMillis() - start) / 1000;
    }

    private void autoSizeColumns(Sheet sheet, int noOfColumns) {
        for (int i = 0; i < noOfColumns; i++) {
            sheet.autoSizeColumn((short) i);
        }
    }

    private Row getOrCreateNextRow(Sheet sheet, int rowNo) {
        Row row;
        if (sheet.getRow(rowNo) != null) {
            row = sheet.getRow(rowNo);
        } else {
            row = sheet.createRow(rowNo);
        }
        return row;
    }

    private CellStyle setCurrencyCellStyle(Workbook workbook) {
        CellStyle currencyStyle = workbook.createCellStyle();
        currencyStyle.setWrapText(true);
        DataFormat df = workbook.createDataFormat();
        currencyStyle.setDataFormat(df.getFormat("#0.00"));
        return currencyStyle;
    }

    private Font getBoldFont(Workbook workbook) {
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeight((short) (10 * 20));
        font.setFontName("Calibri");
        font.setColor(IndexedColors.BLACK.getIndex());
        return font;
    }

    private Font getGenericFont(Workbook workbook) {
        Font font = workbook.createFont();
        font.setFontHeight((short) (10 * 20));
        font.setFontName("Calibri");
        font.setColor(IndexedColors.BLACK.getIndex());
        return font;
    }

    private CellStyle getCenterAlignedCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
        cellStyle.setBorderTop(BorderStyle.NONE);
        cellStyle.setBorderBottom(BorderStyle.NONE);
        cellStyle.setBorderLeft(BorderStyle.NONE);
        cellStyle.setBorderRight(BorderStyle.NONE);
        return cellStyle;
    }

    private CellStyle getLeftAlignedCellStyle(Workbook workbook, Font font) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.LEFT);
        cellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
        cellStyle.setBorderTop(BorderStyle.NONE);
        cellStyle.setBorderBottom(BorderStyle.NONE);
        cellStyle.setBorderLeft(BorderStyle.NONE);
        cellStyle.setBorderRight(BorderStyle.NONE);
        return cellStyle;
    }

}
