package poiexcel.annotation;

import java.lang.annotation.*;

/**
 * 
 * 用于普通类型字段
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelAttribute {

	/**
	 * 导出到Excel中的名字.
	 */
	String name();

	/**
	 * 配置列的名称,对应A,B,C,D....
	 */
	String column();

	String code() default "";

	/**
	 * 提示信息
	 */
	String prompt() default "";

	/**
	 * 设置只能选择不能输入的列内容.
	 */
	String[] combo() default {};

	/**
	 * 是否导出数据,应对需求:有时我们需要导出一份模板,这是标题需要但内容需要用户手工填写.
	 */
	boolean isExport() default true;

	int width() default 10;

	/**如果该字段支持图片格式，其值应为图片路径（支持多图片，图片路径以逗号‘，’分隔）*/
	boolean isImageFormat() default false;

	boolean isLink() default  false ;
}