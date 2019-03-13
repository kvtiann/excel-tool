package poiexcel.annotation;

import java.lang.annotation.*;


/**
 * 
* 用于确定数据的唯一性
*
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelID {

	/**
	 * 
	* 默认属性
	* @return  String 返回类型  
	* @throws
	 */
	String value() default "";
}
