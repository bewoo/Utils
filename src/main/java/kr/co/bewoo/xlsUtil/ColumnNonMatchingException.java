package kr.co.bewoo.xlsUtil;
public class ColumnNonMatchingException extends RuntimeException {

    public ColumnNonMatchingException() {
        super("Columns와 Titles의 개수가 일치하지 않습니다.");
        // TODO Auto-generated constructor stub
    }
}
