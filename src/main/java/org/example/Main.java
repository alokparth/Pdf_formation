package org.example;

import com.itextpdf.io.image.ImageData;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException {
        String psth="invoice.pdf";
        PdfWriter writer =new PdfWriter(psth);
        PdfDocument doc =new PdfDocument(writer);
        Document document=new Document(doc);
        document.add(new Paragraph("Hello"));
        FileInputStream fileInputStream = new FileInputStream("C:\\Users\\UNDERWORLD\\Postman\\files\\Data.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);
        table(sheet,document);
        document.add(new Paragraph("Image Incoming Alert!"));
        String  imgPath="C:\\Users\\UNDERWORLD\\Desktop\\duck.jpeg";
        img(document,imgPath);
        String img2 ="C:\\Users\\UNDERWORLD\\Desktop\\ducks.jpg";
        img(document,img2);
        document.close();
    }

    public static void table(XSSFSheet sheet,Document document){
        Object[] columnNames = {"Id","First Name","Sur Name","Hobby","Profession","Age"};
        Table table=new Table(sheet.getRow(0).getPhysicalNumberOfCells());
        for(int i=0;i<6;i++){
            table.addCell((String) columnNames[i]);
        }
        int rowIndex = 0;
        for (Row row : sheet) {
            List<String> cols=new ArrayList<>();
            if (rowIndex == 0) {
                rowIndex++;
                continue;
            }
            Cell cell = row.getCell(0);
            cols.add(String.valueOf(cell.getNumericCellValue()));
            cell = row.getCell(1);
            cols.add(cell.getStringCellValue());
            cell = row.getCell(2);
            cols.add(cell.getStringCellValue());
            cell = row.getCell(3);
            cols.add(cell.getStringCellValue());
            cell = row.getCell(4);
            cols.add(cell.getStringCellValue());
            cell = row.getCell(5);
            cols.add(String.valueOf(cell.getNumericCellValue()));
            add(cols,table);
        }
        System.out.println("njnni");
        System.out.println(table);
        document.add(table);
    }
    public static void add(List<String> ls, Table table){
        for(int i=0;i<6;i++){
            table.addCell((String) ls.get(i));
        }
    }

    public static void img(Document doc,String imgPath) throws  MalformedURLException {
        ImageData imageData = ImageDataFactory.create(imgPath);
        Image img = new Image(imageData);
        doc.add(img);
    }
}