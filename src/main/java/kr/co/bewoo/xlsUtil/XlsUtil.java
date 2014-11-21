package kr.co.bewoo.xlsUtil;

import java.io.IOException;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import kr.co.bewoo.cmmUtil.StringUtil;
import kr.co.bewoo.reflection.invoke;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;

public class XlsUtil {
    
    /**
     * 타이틀
     */
    public final static short TITLE = 1;
    
    /**
     * 왼쪽정렬
     */
    public final static short LEFT = 2;
    
    /**
     * 가운데 정렬
     */
    public final static short CENTER = 3;
    
    /**
     * 오른쪽 정렬
     */
    public final static short RIGHT = 4;
    
    /**
     * 숫자
     */
    public final static short NUMBER = 5;
    
    /**
     * title HSSFCellStyle
     */
    private HSSFCellStyle title;
    
    /**
     * left HSSFCellStyle
     */
    private HSSFCellStyle lAlign;
    
    /**
     * right HSSFCellStyle
     */
    private HSSFCellStyle rAlign;
    
    /**
     * center HSSFCellStyle
     */
    private HSSFCellStyle cAlign;
    
    /**
     * 타이틀 height, default 40
     */
    private int titleHeight = 40;
    
    /**
     * 컬럼들 각각의 너비
     */
    private int[] columnWidths;
    
    /**
     * 스프레드시트에 작성될 각각의 타이틀
     */
    private String[] titles;
    
    /**
     * invoke될 properties 명
     */
    private String[] properties;
    
    /**
     * 정렬 정보
     */
    private short[] alignInfo;
    
    /**
     * 시트 이름, default = 'sheet1'
     */
    private String sheetName = "sheet1";
    
    /**
     * 스프레드시트 파일명, default = 'EXCEL_'+ yyyyMMdd
     */
    private String documentTitle = "xlsFile";
    
    /**
     * @param columnWidths the columnWidths to set
     */
    public void setColumnWidths(int[] columnWidths) {
        this.columnWidths = columnWidths;
    }
    /**
     * @param titles the titles to set
     */
    public void setTitles(String[] titles) {
        this.titles = titles;
    }
    /**
     * @param titleHeight the titleHeight to set
     */
    public void setTitleHeight(int titleHeight) {
        this.titleHeight = titleHeight;
    }
    /**
     * @param sheetName the sheetName to set
     */
    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }
    /**
     * @param documentTitle the documentTitle to set
     */
    public void setDocumentTitle(String documentTitle) {
        this.documentTitle = documentTitle;
    }
    /**
     * @param properties the properties to set
     */
    public void setProperties(String[] properties) {
        this.properties = properties;
    }
    /**
     * @param alignInfo the alignInfo to set
     */
    public void setAlignInfo(short[] alignInfo) {
        this.alignInfo = alignInfo;
    }
    
    
    /**
     * 엑셀다운로드
     * @param response
     * @throws Exception
     */
    public void XlsDownload(HttpServletResponse response, List<?> list)
            throws Exception {
        int len = columnWidths.length;
        if (columnWidths.length != len && titles.length != len
                && alignInfo.length != len && properties.length != len) {
            throw new ColumnNonMatchingException();
        }

        HSSFWorkbook workBook = createHSSF();
        settingCellInfo(workBook, list);
        settingFileInfo(response, workBook);
    }
   
    /**
     * 스프레드시트 문서, 시트, 폰트, 타이틀, 타이틀 높이, 스타일 객체 생성
     * @return 기본스타일 및 폰트가 적용된 HSSFWorkbook 객체
     */
    private HSSFWorkbook createHSSF() {
        HSSFWorkbook workBook = new HSSFWorkbook();
        HSSFSheet sheet = workBook.createSheet(this.sheetName); 
        HSSFFont font = workBook.createFont();
        HSSFRow row = sheet.createRow(0);
        row.setHeightInPoints(this.titleHeight);
        
        title  = setCellStyle(workBook.createCellStyle(), font, TITLE);
        lAlign = setCellStyle(workBook.createCellStyle(), font, LEFT);
        rAlign = setCellStyle(workBook.createCellStyle(), font, RIGHT);
        cAlign = setCellStyle(workBook.createCellStyle(), font, CENTER);
        
        for(int i =0; i<columnWidths.length; i++) {
            sheet.setColumnWidth(i, 256 * columnWidths[i]);
            createCell(row, title, i, titles[i]);
        }
        
        return workBook;
    }
    
    /**
     * 질의 결과를 전달받아 데이터 작성
     * Map과 Property를 가지고 있는 Java Bean형태의 객체만 처리가능
     * @param workBook
     * @param list
     */
    private void settingCellInfo(HSSFWorkbook workBook, List<?> list) {
        HSSFSheet sheet = workBook.getSheet(this.sheetName);
        HSSFRow row = null;
        HSSFCellStyle style = null;
        Object obj = null;
        for(int i =0; i<list.size(); i++) {
            row = sheet.createRow(i+1);
            for(int j =0; j < properties.length; j++) {
                switch (alignInfo[j]) {
                case LEFT:
                    style = lAlign;
                    break;
                case RIGHT:
                    style = rAlign;
                    break;
                case CENTER:
                    style = cAlign;
                    break;
                default:
                    style = cAlign;
                    break;
                }
                if(list.get(i) instanceof Map) {
                    obj = invoke.invokeByMethodName(list.get(i), "get", properties[j]);
                } else {
                    obj = invoke.invokeByMethodName(list.get(i), "get"+StringUtil.capitalize(properties[j]));
                }
                
                createCell(row, style, j, (obj == null ? "" : obj));
            }
        }
    }
    
    /**
     * 파일 다운로드를 위한 설정
     * @param response
     * @param workBook
     * @throws IOException
     */
    private void settingFileInfo(HttpServletResponse response, HSSFWorkbook workBook) throws IOException {
        String documentTitle = URLEncoder.encode(this.documentTitle,"UTF-8");
        response.setContentType("Content-type: application/octet-stream;charset=UTF-8;");
        response.setHeader("Content-Disposition", "attachment; filename=" + documentTitle + ".xls" );
        response.setHeader("Content-Transfer-Encoding", "binary;");
        response.setHeader("Pragma", "no-cache;");
        response.setHeader("Expires", "-1;");
        workBook.write(response.getOutputStream());
    }
    
    
    /**
     * 기본 셀 생성
     * @param row
     * @param style
     * @param column
     * @param value
     */
    private static void createCell(HSSFRow row, HSSFCellStyle style, int column, Object value) {
        boolean isString  = (value instanceof String);
        HSSFCell cell = row.createCell(column, (isString ? Cell.CELL_TYPE_STRING : Cell.CELL_TYPE_NUMERIC));
        if(isString) {
            cell.setCellValue(new HSSFRichTextString((String)value));
        } else {
            cell.setCellValue(((BigDecimal)value).doubleValue());
        }
        cell.setCellStyle(style);
    }
    
    /**
     * 옵션에 따른 셀 스타일 셋팅
     * @param style
     * @param font
     * @param option
     * @return
     */
    private static HSSFCellStyle setCellStyle(HSSFCellStyle style, HSSFFont font, int option) {

        // 기본 스타일
        style.setWrapText(true);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

        switch (option) {
        case TITLE:
            font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
            style.setFont(font);
            style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            style.setFillForegroundColor(HSSFColor.YELLOW.index);
            style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            break;
        case LEFT:
            style.setAlignment(HSSFCellStyle.ALIGN_LEFT);
            break;
        case CENTER:
            style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            break;
        case RIGHT:
            style.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
            break;
        case NUMBER:
            style.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
            break;
        default:
            style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            break;
        }
        return style;
    }
}
