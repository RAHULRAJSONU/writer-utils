package io.github.rahulrajsonu.writerutils.example;

import io.github.rahulrajsonu.writerutils.annotation.XlsxCompositeField;
import io.github.rahulrajsonu.writerutils.annotation.XlsxSheet;
import io.github.rahulrajsonu.writerutils.annotation.XlsxSingleField;
import io.github.rahulrajsonu.writerutils.service.Exportable;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.sql.Date;
import java.util.List;

@Getter
@NoArgsConstructor
@AllArgsConstructor
@Setter
@XlsxSheet(value = "Users")
public class XlsxUser implements Exportable{

    @XlsxSingleField(columnIndex = 0, columnHeader = "NAME")
    private String name;
    @XlsxSingleField(columnIndex = 1, columnHeader = "GENDER")
    private String gender;
    @XlsxSingleField(columnIndex = 2, columnHeader = "AGE")
    private Integer age;
    @XlsxSingleField(columnIndex = 3, columnHeader = "BMI_VALUE")
    private Double bmiValue;
    @XlsxSingleField(columnIndex = 4, columnHeader = "DOB")
    private Date dob;
    @XlsxSingleField(columnIndex = 5, columnHeader = "IS_OVER_WEIGHT")
    private Boolean isOverweight;
    @XlsxSingleField(columnIndex = 6, columnHeader = "ACTIVITIES")
    private List<String> activities;
    @XlsxCompositeField(from = 7, to = 8)
    private List<XlsxDietPlan> plans;

    @Getter
    @AllArgsConstructor
    @NoArgsConstructor
    @Setter
    public static class XlsxDietPlan implements Exportable {
        @XlsxSingleField(columnIndex = 7, columnHeader = "MEAL_NAME")
        private String mealName;
        @XlsxSingleField(columnIndex = 8, columnHeader = "CALORIES")
        private Double calories;
    }

}