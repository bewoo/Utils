

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

/**  
 * @Project		: GSPowerWeb
 * @Class Name 	: kr.co.gspower.util.invoke.java
 * @Autor		: OOO
 * @Description : 
 * @Modification Information  
 * @
 * @  수정일          수정자              수정내용
 * @ -------------   ---------   -------------------------------
 * @ 2014. 11. 5.     OOO                최초생성
 * 
 * 
 */
public class invoke {
    
    /**
     * 
     * @param obj
     * @param methodName
     * @param objList
     * @return
     */
    public static Object invokeByMethodName(Object obj, String methodName, Object[] objList) {
        Method[] methods = obj.getClass().getMethods();

        for (int i = 0; i < methods.length; i++) {
            if (methods[i].getName().equals(methodName)) {
                try {
                    if (methods[i].getReturnType().getName().equals("void")) {
                        methods[i].invoke(obj, objList);
                    } else {
                        return methods[i].invoke(obj, objList);
                    }
                } catch (IllegalAccessException lae) {
                    System.out.println("LAE : " + lae.getMessage());
                } catch (InvocationTargetException ite) {
                    System.out.println("ITE : " + ite.getMessage());
                }
            }
        }
        return null;
    }

}
