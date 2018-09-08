package com.iqhr;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.lang3.ArrayUtils;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.Reader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.function.Consumer;

public class Main {
    public static void main(String[] args) throws Exception {
        //args = getTestArgs();
        Date date = new Date();
        if(args.length == 0) {
            System.out.println("File location was not provided.");
            return;
        }
        String path = args[0];
        args = ArrayUtils.remove(args, 0);
        String frontNameParam = null;
        String endNameParam = null;
        if(args.length>=2) {
            frontNameParam = args[0];
            args = ArrayUtils.remove(args, 0);
        }
        else {
            frontNameParam = String.valueOf(LocalDate.now().getMonthValue()) + String.valueOf(LocalDate.now().getYear());
        }
        if (args.length >= 2) {
            endNameParam = args[0];
            args = ArrayUtils.remove(args, 0);
        } else {
            endNameParam = String.valueOf(LocalDate.now().getMonthValue()) + String.valueOf(LocalDate.now().getYear());
        }
        List<String> sortColumns = Arrays.asList(args);
        Path pathToProcess = Paths.get(path);
        if(Files.notExists(pathToProcess)) {
            System.out.println("Could not find file.");
            return;
        }

        Service.processPath(pathToProcess, frontNameParam, endNameParam, sortColumns);
        System.out.println("Processing finished!");
        System.out.println("Processing took " + (new Date().getTime()-date.getTime())/1000 + " seconds");
    }

    private static String[] getTestArgs() {
        //return new String[] {"C:\\work\\personal_projects\\excelautomation\\src\\main\\resources\\", "082012","companie","nume"};
        return new String[]{"C:\\work\\pt muma\\rapoarte_mari"};
    }


}
