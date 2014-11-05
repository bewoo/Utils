/**  
 * @Project		: GSPowerWeb
 * @Class Name 	: kr.co.gspower.util.ColumnNonMatchingException.java
 * @Autor		: OOO
 * @Description : 
 * @Modification Information  
 * @
 * @  수정일          수정자              수정내용
 * @ -------------   ---------   -------------------------------
 * @ 2014. 11. 4.     OOO                최초생성
 * 
 * 
 */


public class ColumnNonMatchingException extends RuntimeException {

    public ColumnNonMatchingException() {
        super("Columns와 Titles의 개수가 일치하지 않습니다.");
        // TODO Auto-generated constructor stub
    }
}
