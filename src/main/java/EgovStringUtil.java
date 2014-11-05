import org.apache.commons.lang.text.StrBuilder;

public class EgovStringUtil {
    /**
     * Character 객체에서 첫번째 char만 대문자로 바꾸고 두번째 char부터는 append를 해주어 String을 생성한다.
     * 
     * @param str
     * @return
     */
    public static String capitalize(String str) {
        int strLen;
        if (str == null || (strLen = str.length()) == 0) {
            return str;
        }
        return new StrBuilder(strLen)
                .append(Character.toTitleCase(str.charAt(0)))
                .append(str.substring(1)).toString();
    }
}
