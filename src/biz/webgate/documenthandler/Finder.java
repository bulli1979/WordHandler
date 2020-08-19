package biz.webgate.documenthandler;
import java.util.ArrayList;
import java.util.List;

import org.docx4j.TraversalUtil.CallbackImpl;


public class Finder extends CallbackImpl {

    protected Class<?> typeToFind;

    protected Finder(Class<?> typeToFind) {
        this.typeToFind = typeToFind;
    }

      public List<Object> results = new ArrayList<Object>(); 

      public List<Object> apply(Object o) {    	  
          if (o.getClass().equals(typeToFind)) {
              results.add(o);
          }
          return null;
      }
}
