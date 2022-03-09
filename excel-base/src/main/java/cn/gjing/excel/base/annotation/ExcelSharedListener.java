package cn.gjing.excel.base.annotation;

import org.springframework.stereotype.Component;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Declare the listener to be a shared listener and be managed by the Spring IOC container.
 * Listeners are referenced by all imports and exports, so it is best not to use global variables because of thread-safety issues.
 *
 * @author Gjing
 **/
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
@Component
public @interface ExcelSharedListener {
    /**
     * Deny sharing to some Excel entities
     *
     * @return ignore excel entities
     */
    Class<?>[] ignore();
}
