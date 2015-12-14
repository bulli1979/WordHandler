package biz.webgate.documenthandler;

import java.util.List;

import org.junit.Test;
public class ActiveXElementTest {
	@Test
	public void testActiveXElement(){
		
		WordHandler handler = new WordHandler();
		handler.read();
		List<ElementType> cbList = handler.getCBs();
		List<ElementType> tList = handler.getTextList();
		System.out.println(cbList.size() + " : " + tList.size());
	}
}
