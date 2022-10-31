//package io.github.rahulrajsonu.writerutils.factory;
//
//import org.apache.poi.ss.usermodel.Workbook;
//import org.slf4j.Logger;
//import org.slf4j.LoggerFactory;
//
//import java.util.List;
//
//public class ExcelUtils {
//
//    private static final Logger logger = LoggerFactory.getLogger(ExcelUtils.class);
//
//    public static <T> Workbook generateExcel(String filter, ExcelTemplate template, List<? extends T> dataList) throws Exception {
//        List<String> fields = analyzeFilter(filter);
//        if(logger.isDebugEnabled()){
//        	for(String field:fields){
//            	logger.debug("field want to be set in excel is "+field);
//            }
//        }
//
//        List<ExcelTemplate> excelTemplates = buildTemplateList(fields, template);
//        List<List<Object>> values = buildValueList(dataList, fields);
//        Workbook basicWorkbook = createBasicWorkbook(excelTemplates, values);
//        return basicWorkbook;
//    }
//
//    /**
//     * 根据templete生成默认格式的excel文件
//     *
//     * @param templateList
//     * @param cellValues
//     * @return
//     * @author dan.shan
//     * @since Jul 8, 2013 11:53:55 AM
//     */
//    private static Workbook createBasicWorkbook(List<ExcelTemplate> templateList, List<List<Object>> cellValues) {
//
//        Workbook workbook = new XSSFWorkbook();
//        Sheet sheet = workbook.createSheet("sheet");
//        ExcelTemplate template;
//
//        CellStyle style = workbook.createCellStyle();            // 这里有颜色列表 http://jlcon.iteye.com/blog/1122538
//
//        // generate excel title
//        Row row = sheet.createRow(0);
//        for (int i = 0; i < templateList.size(); i++) {
//            Cell cell = row.createCell(i, XSSFCell.CELL_TYPE_STRING);
//            template = templateList.get(i);
//            cell.setCellValue(template.getTitle());
//            cell.setCellStyle(style);
//        }
//
//        DataFormat dataFormat = workbook.createDataFormat();
//
//        CellStyle stylePercent = workbook.createCellStyle();
//        stylePercent.setDataFormat(dataFormat.getFormat("0.00%"));
//
//        CellStyle styleTwoDecimalPlaces = workbook.createCellStyle();
//        styleTwoDecimalPlaces.setDataFormat(dataFormat.getFormat("#,##0.00"));
//
//        CellStyle styleThousandPlace = workbook.createCellStyle();
//        styleThousandPlace.setDataFormat(dataFormat.getFormat("#,##0"));
//
//        // generate excel date content
//        if (cellValues != null) {
//            for (List<Object> line : cellValues) {
//                row = sheet.createRow(sheet.getLastRowNum() + 1);
//                for (int i = 0; i < line.size(); i++) {
//                    Cell cell = row.createCell(i);
//                    ExcelTemplate excelTemplate = templateList.get(i);
//
//                    Object valueObject = line.get(i);
//
//                    // Null gets blank, otherwise gets toString() value
//                    Integer cellType;
//                    String cellValue;
//                    if (valueObject == null) {
//                        cellType = ExcelCellType.BLANK;
//                        cellValue = "";
//                    } else {
//                        cellType = excelTemplate.getCellType();
//                        cellValue = valueObject.toString();
//                    }
//
//                    // inject value as different type
//                    if (cellType == ExcelCellType.PERCENTAGE) {
//                        cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
//                        cell.setCellValue(NumberUtils.toFloat(cellValue));
//                        cell.setCellStyle(stylePercent);
//                    } else if (cellType == ExcelCellType.TWO_DECIMAL_PLACES) {
//                        cell.setCellValue(NumberUtils.toFloat(cellValue));
//                        cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
//                        cell.setCellStyle(styleTwoDecimalPlaces);
//                    } else if (cellType == ExcelCellType.THOUSAND_PLACE) {
//                        cell.setCellValue(NumberUtils.toFloat(cellValue));
//                        cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
//                        cell.setCellStyle(styleThousandPlace);
//                    } else if (cellType == ExcelCellType.INT) {
//                        cell.setCellValue(NumberUtils.toInt(cellValue));
//                        cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
//                        cell.setCellStyle(style);
//                    } else if (cellType == ExcelCellType.DATE) {
//                        cell.setCellValue(cellValue);
//                    }else if(cellType == ExcelCellType.BLANK){
//                        cell.setCellValue(cellValue);
//                    } else {
//                        cell.setCellStyle(style);
//                        cell.setCellType(cellType);
//                        cell.setCellValue(cellValue);
//                    }
//                }
//            }
//        }
//
//        for (int i = 0; i < templateList.size(); i++) {
//            sheet.autoSizeColumn(i);
//        }
//
//        return workbook;
//    }
//
//    /**
//     * analyze filter string, set up a list
//     *
//     * @param filter
//     * @return
//     */
//    private static List<String> analyzeFilter(String filter) {
//        if (StringUtils.isBlank(filter)) {
//            throw new IllegalArgumentException("filter can not be null");
//        }
//
//        String[] fields = filter.split(",");
//        List<String> fieldList = new ArrayList<String>();
//        for (String field : fields) {
//            if (StringUtils.isNotBlank(field) && !fieldList.contains(field)) {
//            	if("undefined".equals(field)){
//            		continue;
//            	}
//                fieldList.add(field);
//            }
//        }
//
//        return fieldList;
//    }
//
//    /**
//     * Based on custom fields, find proper excel templates
//     *
//     * @param fields
//     * @return
//     */
//    private static List<ExcelTemplate> buildTemplateList(List<String> fields, ExcelTemplate template) {
//
//        if (CollectionUtils.isEmpty(fields)) {
//            throw new IllegalArgumentException("Filter can not be null during downloading");
//        }
//
//        List<ExcelTemplate> templateList = new ArrayList<ExcelTemplate>();
//        for (String field : fields) {
//            try {
//                ExcelTemplate excelTemplate = template.valueOf(field);
//                templateList.add(excelTemplate);
//            } catch (IllegalArgumentException e) {
//                logger.error("Unknown field name {}", field);
//                continue;
//            }
//        }
//
//        return templateList;
//    }
//
//    /**
//     * generate excel data
//     *
//     * @param dataList
//     * @param customFields
//     * @return
//     */
//    private static <T> List<List<Object>> buildValueList(List<? extends T> dataList, List<String> customFields) throws Exception {
//        List<List<Object>> result = new ArrayList<List<Object>>();
//        for (T data : dataList) {
//            List<Object> values = new ArrayList<Object>();
//            for (String customField : customFields) {
//                try {
//                    //use util to cache reflection
//                    Object specificValue = PropertyUtils.getProperty(data, customField);
//                    if (specificValue instanceof Date) {
//                        Date date = (Date) specificValue;
//                        String actualValue = DateFormatUtils.format(date, "yyyy/MM/dd");
//                        values.add(actualValue);
//                        continue;
//                    }
//                    values.add(specificValue);
//                } catch (Exception e) {
//                    logger.error("could not get bean property - {}", customField);
//                    throw e;
//                }
//            }
//            result.add(values);
//        }
//
//        return result;
//    }
//}