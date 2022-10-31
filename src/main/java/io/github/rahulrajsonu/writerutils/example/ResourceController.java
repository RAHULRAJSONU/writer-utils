package io.github.rahulrajsonu.writerutils.example;

import io.github.rahulrajsonu.writerutils.config.ExportType;
import io.github.rahulrajsonu.writerutils.factory.SpreadsheetExporter;
import io.github.rahulrajsonu.writerutils.service.Exporter;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.sql.Date;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;

@RestController
@RequestMapping("/download")
public class ResourceController {

    @PostMapping()
    public ResponseEntity<StreamingResponseBody> downloadFile(@RequestBody List<XlsxUser> user,
                                                              HttpServletResponse response) {
        response.setHeader("Content-Disposition", "attachment; filename=User.xls");
        int BUFFER_SIZE = 1024;

        Exporter exporter = new SpreadsheetExporter(ExportType.SPREADSHEET);

        return ResponseEntity.ok(outputStream -> {
            int bytesRead;
            byte[] buffer = new byte[BUFFER_SIZE];

            //convert the file to InputStream
            InputStream inputStream = new ByteArrayInputStream(exporter.export(user,"User.xls"));
            while ((bytesRead = inputStream.read(buffer)) != -1) {
                outputStream.write(buffer, 0, bytesRead);
                outputStream.flush();
            }
        });

    }

    @GetMapping("/getPI")
    public ResponseEntity<List<XlsxUser>> getPersonalInfo(){
        List<XlsxUser> data = new ArrayList<>();
        for(int i=0; i < 10; i++){
            data.add(new XlsxUser("Person: "+i,i%2==0?"Female":"Male",28,22.7,new Date(1653868800000L),false, Collections.emptyList(), Collections.emptyList()));
        }
        return ResponseEntity.ok(data);
    }

    @GetMapping("/map-to-excel")
    public ResponseEntity<StreamingResponseBody> mapToExcel(HttpServletResponse response) throws InvocationTargetException, NoSuchMethodException, IllegalAccessException {
        List<XlsxUser> data = new ArrayList<>();
        for(int i=0; i < 10; i++){
            data.add(new XlsxUser("Person: "+i,i%2==0?"Female":"Male",28,22.7,new Date(1653868800000L),false, Collections.emptyList(), Collections.emptyList()));
        }
        List<LinkedHashMap<String, String>> records = convertToListOfMap(data);
        response.setHeader("Content-Disposition", "attachment; filename=User.xls");
        int BUFFER_SIZE = 1024;

        Exporter exporter = new SpreadsheetExporter(ExportType.MAP_TO_EXCEL);

        return ResponseEntity.ok(outputStream -> {
            int bytesRead;
            byte[] buffer = new byte[BUFFER_SIZE];

            //convert the file to InputStream
            InputStream inputStream = new ByteArrayInputStream(exporter.export(records,"User.xls"));
            while ((bytesRead = inputStream.read(buffer)) != -1) {
                outputStream.write(buffer, 0, bytesRead);
                outputStream.flush();
            }
        });
    }

    private <T> List<LinkedHashMap<String, String>> convertToListOfMap(List<T> data) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        Class<?> clazz = data.get(0).getClass();
        Field[] fields = clazz.getDeclaredFields();
        List<LinkedHashMap<String, String>> list = new ArrayList<>();
        for (T t : data) {
            LinkedHashMap<String, String> record = new LinkedHashMap<>();
            for (Field field : fields) {
                Method method = getMethod(clazz, field);
                Object val = method.invoke(t, (Object[]) null);
                record.put(field.getName(), String.valueOf(val));
            }
            list.add(record);
        }
        return list;
    }

    private Method getMethod(Class<?> clazz, Field field) throws NoSuchMethodException {
        Method method;
        try {
            method = clazz.getMethod("get" + capitalize(field.getName()));
        } catch (NoSuchMethodException nme) {
            method = clazz.getMethod(field.getName());
        }

        return method;
    }

    private static String capitalize(String s) {
        if (s.length() == 0)
            return s;
        return s.substring(0, 1).toUpperCase() + s.substring(1);
    }
}
