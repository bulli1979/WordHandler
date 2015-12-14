package biz.webgate.documenthandler;
import java.util.ArrayList;
import java.util.List;
import org.docx4j.TraversalUtil;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.R.Sym;
import org.docx4j.wml.SdtBlock;
import org.docx4j.wml.SdtContentBlock;
import org.docx4j.wml.SdtPr;
import org.docx4j.wml.SdtRun;
import org.docx4j.wml.Tag;
import org.docx4j.wml.Text;

public class WordHandler {

	private MainDocumentPart documentPart;
	
	public WordHandler() {
		super();
		// TODO Auto-generated constructor stub
	}

	public void read() {
		try {
			String fileName = "C:\\Users\\Mirko\\Desktop\\svz\\IV201500002_20151105_IMP_URTEIL.docm";
			WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File(fileName));
			this.documentPart = wordMLPackage.getMainDocumentPart();
			//System.out.println(this.documentPart.getXML());
		} catch (Exception e) {
			System.out.println("Error in  readWord " + e);
			e.printStackTrace();
		}
	}
	
	
	public List<ElementType> getCBs(){
		List<ElementType> retList = new ArrayList<ElementType>();
		Finder finder = new Finder(Sym.class);
		new TraversalUtil(documentPart.getContent(), finder);
		for (Object obj : finder.results){
			Sym sym = (Sym)obj;
			boolean checked = chkBox(sym.getChar());
			R parent1 = (R)sym.getParent();
			
			SdtRun parent2 = (SdtRun)parent1.getParent();
			String name = getBoxName(parent2.getSdtPr());
			ElementType cb = new ElementType(name, checked,"",0);
			retList.add(cb);
		}
		finder = null;
		return retList;
	}
	
	
	private String getBoxName(SdtPr pr){
		Tag t = pr.getTag();
		return t.getVal();
	}
	
	private boolean chkBox(String code){
		boolean checked = false;
		if(code != null && code.equals("F054")){
			checked = true;
		}
		return checked;
	} 
	
	public List<ElementType> getTextList(){
		System.out.println("starte");
		List<ElementType> retList = new ArrayList<ElementType>();
		Finder finder = new Finder(Text.class);
		new TraversalUtil(documentPart.getContent(), finder);
		for (Object obj : finder.results){
			Text text = (Text)obj;
			String name="";
			System.out.println(text.getValue());
			R parent1 = (R)text.getParent();
			if(parent1.getParent() instanceof P){
				P parent2 = (P)parent1.getParent();
				if(parent2.getParent() instanceof SdtContentBlock){
					SdtContentBlock parent3 = (SdtContentBlock)parent2.getParent();
					if(parent3.getParent() instanceof SdtBlock){
						SdtBlock parent4 = (SdtBlock)parent3.getParent();
						try{
							name = getBoxName(parent4.getSdtPr());
							ElementType t = new ElementType(name, false,text.getValue(),1);
							retList.add(t);
						}catch(Exception e){
						}
						
					}
				}
			}
		}
		finder = null;
		
		
		
		return retList;
	}
	
	
}