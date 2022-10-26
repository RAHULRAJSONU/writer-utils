package io.github.rahulrajsonu.writerutils.example;

import io.github.rahulrajsonu.writerutils.factory.SpreadsheetExporter;
import io.github.rahulrajsonu.writerutils.service.Exporter;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

@RestController
@RequestMapping("/download")
public class ResourceController {

    @GetMapping()
    public StreamingResponseBody downloadFile(HttpServletResponse response) {
        response.setHeader("Content-Disposition", "attachment; filename=User.xls");
        int BUFFER_SIZE = 1024;

        List<XlsxUser> data = new ArrayList<>();
        for(int i=0; i < 100000; i++){
            data.add(new XlsxUser("Person: "+i,i%2==0?"Female":"Male",28,22.7,false, Collections.emptyList(), Collections.emptyList()));
        }
        Exporter exporter = new SpreadsheetExporter();

        return outputStream -> {
            int bytesRead;
            byte[] buffer = new byte[BUFFER_SIZE];

            //convert the file to InputStream
            InputStream inputStream = new ByteArrayInputStream(exporter.export(data,XlsxUser.class,"User.xls"));
            while ((bytesRead = inputStream.read(buffer)) != -1) {
                outputStream.write(buffer, 0, bytesRead);
            }

            if (inputStream != null) {
                inputStream.close();
            }
        };

    }
}
