import java.io.File;
import java.util.Scanner;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class converter {
    public static void main(String[] args) throws Exception {

        String reg="",studn="",p_1,p_2,c_1,c_2;

        Scanner sc=new Scanner(System.in);
        System.out.print("Enter registration number: ");
        reg=sc.nextLine();
        System.out.print("Enter student's name: ");
        studn=sc.nextLine();
        System.out.print("Enter Presentation 1 grade: ");
        p_1=sc.nextLine().toUpperCase();
        System.out.print("Enter Presentation 2 grade: ");
        p_2=sc.nextLine().toUpperCase();
        System.out.print("Enter Case Study 1 grade: ");
        c_1=sc.nextLine().toUpperCase();
        System.out.print("Enter Case Study 2 grade: ");
        c_2=sc.nextLine().toUpperCase();
        
        WritableWorkbook workbook=Workbook.createWorkbook(new File("FirstWorkBook.xls"));
        WritableSheet sheet= workbook.createSheet("sheet_1", 0);

        //creating cells
        WritableCell cell=new Label(0,0,"Registration no.");
        WritableCell cell_1=new Label(0,1,"Name");
        WritableCell cell_2=new Label(0,2,"Presentation 1 marks");
        WritableCell cell_3=new Label(0,3,"Presentation 2 marks");
        WritableCell cell_4=new Label(0,4,"Case Study 1 marks");
        WritableCell cell_5=new Label(0,5,"Case Study 2 marks");
        WritableCell cell_6=new Label(0,6,"Total marks");

        //convert the grades to marks:
        String m_p_1=converter.gradetomark(p_1);
        String m_p_2=converter.gradetomark(p_2);
        String m_c_1=converter.gradetomark(c_1);
        String m_c_2=converter.gradetomark(c_2);


        //creating marks
        WritableCell ans_cell=new Label(1,0,reg);
        WritableCell ans_cell_1=new Label(1,1,studn);
        WritableCell ans_cell_2=new Label(1,2,m_p_1);
        WritableCell ans_cell_3=new Label(1,3,m_p_2);
        WritableCell ans_cell_4=new Label(1,4,m_c_1);
        WritableCell ans_cell_5=new Label(1,5,m_c_2);
        WritableCell ans_cell_6=new Label(1,6,converter.gettotalmark(m_p_2, m_c_1, m_c_2, m_p_1));

        //adding cells
        sheet.addCell(cell);
        sheet.addCell(cell_1);
        sheet.addCell(cell_2);
        sheet.addCell(cell_3);
        sheet.addCell(cell_4);
        sheet.addCell(cell_5);
        sheet.addCell(cell_6);

        //adding answer cells
        sheet.addCell(ans_cell);
        sheet.addCell(ans_cell_1);
        sheet.addCell(ans_cell_2);
        sheet.addCell(ans_cell_3);
        sheet.addCell(ans_cell_4);
        sheet.addCell(ans_cell_5);
        sheet.addCell(ans_cell_6);

        
        workbook.write();
        workbook.close();
        sc.close();
    }

    static String gradetomark(String i){
        if(i.equals("A+")) return "10";
        else if(i.equals("A")) return "9";
        else if(i.equals("B+")) return "8";
        else if(i.equals("B")) return "7";
        else if(i.equals("C+")) return "6";
        else if(i.equals("C")) return "5";
        else return "0";
    }

    static String gettotalmark(String a,String b,String c,String d){
        return Integer.toString(Integer.parseInt(a)+Integer.parseInt(b)+Integer.parseInt(c)+Integer.parseInt(d));
    }
}

/*

Write a java  code  and takes input data 
1.Reg.No
2.Name
3.Presentation 1 Grade
4.Presentation 2 Grade
5.Case study 1 marks
6.Case study 2 marks
The output should be opened in Ms Excel such that the input data should be displayed and presentation grades should be converted into marks as follows
A+=10
A=9
B+=8
B=7
C+=6
C=5
and including total marks.

 */