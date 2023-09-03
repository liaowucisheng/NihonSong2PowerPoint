package com.liaowucisheng;

import org.apache.poi.xslf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

public class App {
    public static void main(String[] args) {
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            XSLFSlide slide = ppt.createSlide();
            XSLFTextShape title = slide.createTextBox();
            XSLFTextParagraph paragraph = title.addNewTextParagraph();
            XSLFTextRun run = paragraph.addNewTextRun();
            run.setText("Hello, World!");

            FileOutputStream out = new FileOutputStream("src/main/resources/presentation.pptx");
            ppt.write(out);
            out.close();
            System.out.println("Slide added successfully!");
        } catch (IOException e) {
            System.out.println("Error adding slide: " + e.getMessage());
        }
    }
}
