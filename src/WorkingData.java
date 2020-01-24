import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;

public class WorkingData {
    private ArrayList<ArrayList<String>> list = null;
    private ArrayList<String> dateList = null;
    private ArrayList<String> weekList = null;

    public WorkingData()
    {
        list = new ArrayList<ArrayList<String>>();
        dateList = new ArrayList<String>();
        weekList = new ArrayList<String>();
    }

    public ArrayList<ArrayList<String>> getList() {
        return list;
    }

    public ArrayList<String> getWeekList() {
        return weekList;
    }

    public ArrayList<String> getDateList() {
        return dateList;
    }

    public void createAttribute(XSSFSheet sheet, int rowIndex) //맨 위 속성값 생성
    {
        XSSFRow row = sheet.getRow(rowIndex);
        if (row != null) {
            int column = row.getPhysicalNumberOfCells();
            if(rowIndex == 0)
            {
                dateList = saveColumn(column, row);
            }
            else if (rowIndex == 1)
            {
                weekList = saveColumn(column, row);
            }
        }
    }

    public void createPersonData(XSSFSheet sheet, int rowIndex)
    {
        for(int x= 2;x<rowIndex;x++) // 앞 2개는 속성 데이터
        {
            XSSFRow xssfRow = sheet.getRow(x);
            if(xssfRow !=null)
            {
                int columnIndex = xssfRow.getPhysicalNumberOfCells();
                list.add(saveColumn(columnIndex, xssfRow));
            }
        }
    }

    public ArrayList<String> saveColumn (int columnIndex, XSSFRow row) //각 열에 대한 행 저장
    {
        ArrayList<String> returnList = new ArrayList<String>();
        for(int x = 0;x<columnIndex;x++)
        {
            XSSFCell xssfcell = row.getCell(x);
            String cellData = dataRead(xssfcell);
            if(cellData !=null)
            {
                returnList.add(cellData);
                System.out.print(cellData+ "\t");
            }
        }
        System.out.println("");
        return returnList;
    }

    public String dataRead(XSSFCell cell) //셀에 있는 데이터를 String 양식으로 가져오기
    {
        String data =  null;
        try
        {
            if(cell.getCellType() == CellType.NUMERIC)
            {
                data = Integer.toString((int)cell.getNumericCellValue());
            }
            else
            {
                data = cell.getStringCellValue();
            }
        }
        catch (NullPointerException e)
        {
            data = "";
            e.printStackTrace();
        }
        return data;
    }

}
