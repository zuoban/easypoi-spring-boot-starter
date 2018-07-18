package com.zuoban.easypoi;

import org.springframework.boot.autoconfigure.AutoConfigureAfter;
import org.springframework.boot.autoconfigure.web.WebMvcAutoConfiguration;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.view.BeanNameViewResolver;

/**
 * @author wangjinqiang
 * @date 2018-07-18
 */

@Configuration
@AutoConfigureAfter({WebMvcAutoConfiguration.class})
public class EasyPoiAutoConfiguration {

    @Bean
    public BeanNameViewResolver beanNameViewResolver() {
        BeanNameViewResolver resolver = new BeanNameViewResolver();
        resolver.setOrder(10);
        return resolver;
    }
}
