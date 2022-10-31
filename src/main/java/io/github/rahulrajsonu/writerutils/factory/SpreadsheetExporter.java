package io.github.rahulrajsonu.writerutils.factory;

import io.github.rahulrajsonu.writerutils.config.ExportType;
import io.github.rahulrajsonu.writerutils.service.Exporter;
import io.github.rahulrajsonu.writerutils.service.Writer;

public class SpreadsheetExporter extends Exporter {

    public SpreadsheetExporter(String exportType) {
        super(exportType);
    }
    @Override
    public Writer createWriter(String exportType) {
        if(exportType.equals(ExportType.SPREADSHEET)) {
            return new SpreadsheetWriter();
        } else if (exportType.equals(ExportType.MAP_TO_EXCEL)) {
            return new MapToSpreadsheetWriter();
        }
        throw new IllegalArgumentException("No such writer: "+exportType+" found.");
    }
}
