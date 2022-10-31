package io.github.rahulrajsonu.writerutils.factory;

import io.github.rahulrajsonu.writerutils.service.Writer;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayOutputStream;
import java.sql.Date;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;

public class MapToSpreadsheetWriter implements Writer {

    private static final Logger logger = LoggerFactory.getLogger(MapToSpreadsheetWriter.class);
    @Override
    public <T> byte[] write(List<T> data, String sheetName) {

        long start = System.currentTimeMillis();
        if(!data.isEmpty()) {
            try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream();) {
                ArrayList<LinkedHashMap<String, String>> records = (ArrayList<LinkedHashMap<String, String>>)(ArrayList<?>)data;
                // setting up the basic styles for the workbook
                CellStyle currencyStyle = setCurrencyCellStyle(workbook);
                CellStyle centerAlignedStyle = getCenterAlignedCellStyle(workbook);
                CreationHelper createHelper = workbook.getCreationHelper();
                CellStyle dateCellStyle = workbook.createCellStyle();
                dateCellStyle.setDataFormat(
                        createHelper.createDataFormat().getFormat("yyyy/MM/dd"));
                Map<String, CellStyle> styleMap = new ConcurrentHashMap<>();
                styleMap.put("CURRENCY_STYLE",currencyStyle);
                styleMap.put("CENTER_ALIGNMENT_STYLE",centerAlignedStyle);
                styleMap.put("DATE_STYLE",dateCellStyle);
                Sheet sheet = workbook.createSheet(sheetName);
                Optional<Set<String>> oHeaders = extractHeaders(records);
                Set<String> headers = oHeaders.orElseThrow(() -> new RuntimeException("Headers not found for the : Map to Excel conversion operation."));
                int rowNum = 0;
                Row headerRow = sheet.createRow(rowNum++);
                int headerCellNum= 0;
                for (String header : headers) {
                    Cell cell = headerRow.createCell(headerCellNum++);
                    cell.setCellValue(header);
                }
                for (LinkedHashMap<String, String> map : records) {
                    Row currentRow = sheet.createRow(rowNum);
                    int cellNum=0;
                    for (String header: headers) {
                        Cell cell = currentRow.createCell(cellNum);
                        if (map.containsKey(header)) {
                            setCellValue(cell,map.get(header),styleMap,workbook);
                        }
                        cellNum++;
                    }
                    rowNum++;
                }
                workbook.write(out);
                logger.info("Xls file generated in [{}] seconds", processTime(start));
                return out.toByteArray();
            } catch (Exception e) {
                logger.info("Xls file write failed", e);
                throw new RuntimeException("Xls file write failed, error: " + e.getMessage());
            }
        }else {
            logger.info("No data found to convert to excel.");
            return new byte[0];
        }
    }

    private Optional<Set<String>> extractHeaders(ArrayList<LinkedHashMap<String, String>> records) {
        return records.stream()
                .map(LinkedHashMap::keySet).findFirst();
    }

    private void setCellValue(Cell cell, Object objValue, Map<String, CellStyle> styleMap, Workbook workbook) {

        Hyperlink link = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);

        if (objValue != null) {
            if (objValue instanceof String) {
                String cellValue = (String) objValue;
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
                cell.setCellStyle(styleMap.get("CURRENCY_STYLE"));
                cell.setCellValue(cellValue);
            } else if (objValue instanceof Boolean) {
                cell.setCellStyle(styleMap.get("CENTER_ALIGNMENT_STYLE"));
                if (objValue.equals(true)) {
                    cell.setCellValue(1);
                } else {
                    cell.setCellValue(0);
                }
            } else if(objValue instanceof Date){
                Date date = (Date)objValue;
                cell.setCellValue(date);
                cell.setCellStyle(styleMap.get("DATE_STYLE"));
            }
        }
    }

    private static String capitalize(String s) {
        if (s.length() == 0)
            return s;
        return s.substring(0, 1).toUpperCase() + s.substring(1);
    }

    private long processTime(long start) {
        return (System.currentTimeMillis() - start) / 1000;
    }

    private CellStyle setCurrencyCellStyle(Workbook workbook) {
        CellStyle currencyStyle = workbook.createCellStyle();
        currencyStyle.setWrapText(true);
        DataFormat df = workbook.createDataFormat();
        currencyStyle.setDataFormat(df.getFormat("#0.00"));
        return currencyStyle;
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

}
