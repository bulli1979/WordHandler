package biz.webgate.documenthandler;

import java.util.ArrayList;
import java.util.List;

import org.apache.log4j.Logger;
import org.docx4j.TraversalUtil;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.w14.CTSdtCheckbox;
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
	private boolean showSysOut = false;
	private boolean debug = false;
	
	final static Logger LOG = Logger.getLogger(WordHandler.class);
	
	private String version = "20173011";
	
	public WordHandler() {
		super();
	}

	public void read() {
		read("");
	}
	
	public String getVersion() {
		return version;
	}

	public boolean isDebug() {
		return debug;
	}
	public void setDebug(boolean debug) {
		this.debug = debug;
	}

	public void read(String fileName) {
		ClassLoader cl = Thread.currentThread().getContextClassLoader();
		try {
			Thread.currentThread().setContextClassLoader(WordHandler.class.getClassLoader());
			LOG.info("changed ClassLoader");
			
			if (debug) System.out.println("Start WordHandler Read");
			
			WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File(fileName));
			this.documentPart = wordMLPackage.getMainDocumentPart();
			
			if (debug) {
				LOG.info("********************XML for " + fileName + ":");
				LOG.info(this.documentPart.getXML());
			}
		} catch (Exception e) {
			LOG.error(e);
			e.printStackTrace();
		} finally {
			Thread.currentThread().setContextClassLoader(cl);
			LOG.info("reset ClassLoader");
		}
	}
	
	
	public List<ElementType> getCBs() {
		List<ElementType> retList = new ArrayList<ElementType>();
		
		try {
			Finder finderSdtRun = new Finder(SdtRun.class);
			new TraversalUtil(this.documentPart.getContent(), finderSdtRun);
			String tagName = "";
			
			for (Object obj : finderSdtRun.results) {
				SdtRun stdRun = (SdtRun) obj;
				SdtPr sdtPr = stdRun.getSdtPr();
				if (sdtPr != null) {
					tagName = sdtPr.getTag().getVal();
					System.out.println("checking " + tagName);
					
					CTSdtCheckbox cbx = (CTSdtCheckbox) sdtPr.getByClass(CTSdtCheckbox.class);
					if (cbx != null) {
						System.out.println(tagName + " - getChecked: " + cbx.getChecked().getVal());
						System.out.println(tagName + " - getCheckedState: " + cbx.getCheckedState().getVal());
						System.out.println(tagName + " - getUncheckedState: " + cbx.getUncheckedState().getVal());
												
						if (cbx.getChecked().getVal() != null) {
							if (cbx.getChecked().getVal().equals("1")) {
								ElementType cb = new ElementType(tagName, true,"",0);
								retList.add(cb);
							} else if (cbx.getChecked().getVal().equals("0")) {
								ElementType cb = new ElementType(tagName, false,"",0);
								retList.add(cb);
							}
						} else {
							System.out.println(tagName + " - getChecked().getVal() is NULL");
						}
					}
				}
			}
			finderSdtRun = null;
		} catch (Exception e) {
			System.out.println("Error in WordHandler.getCBs: " + e);
			e.printStackTrace();
		}
		
		return retList;
	}
	
	// To be removed
	/*
	public List<ElementType> getCBs_old() {
		List<ElementType> retList = new ArrayList<ElementType>();
		try {			
			Finder finder = new Finder(Sym.class);
			new TraversalUtil(documentPart.getContent(), finder);
			for (Object obj : finder.results){
				Sym sym = (Sym)obj;
				System.out.println(sym.getChar() + " > " + sym.getFont());
				
				boolean checked = chkBox(sym.getChar());
				R parent1 = (R)sym.getParent();
				Object parent = parent1.getParent();
				if (parent instanceof SdtRun) {
					SdtRun parent2 = (SdtRun) parent;
					String name = getBoxName(parent2.getSdtPr());
					
					ElementType cb = new ElementType(name, checked,"",0);
					retList.add(cb);
				}
			}
			finder = null;
		} catch (Exception e) {
			System.out.println("Error in WordHandler.getCBs: " + e);
			e.printStackTrace();
		}
		
		return retList;
	}
	*/
	
	
	private String getBoxName(SdtPr pr){
		Tag t = pr.getTag();
		return t.getVal();
	}
	
	/*
	private boolean chkBox(String code){
		//checked		<w:sym w:font="Wingdings 2" w:char="F053"/>		evt. <w:sym w:font="Wingdings 2" w:char="F054"/>
		//				in Word: Wingdings 2 > 83
		//not checked:	<w:sym w:font="Wingdings 2" w:char="F0A3"/>		evt. <w:sym w:font="Wingdings" w:char="F0A8"/>
		//				in Word: Wingdings 2 > 163
		
		//<w:sym w:font="Wingdings" w:char="F06F"/>
		boolean checked = false;
		if(code != null && (code.equals("F053") || code.equals("0053") || code.equals("F054") || code.equals("0054"))) {
			checked = true;
		} else if(!code.equals("F0A3") && !code.equals("00A3") && !code.equals("F0A8") && !code.equals("00A8")) {
			System.out.println("Invalid code: " + code + " - Check Symbol in Word Document");
		}
		return checked;
	}
	*/ 
	
	public List<ElementType> getTextList(){
		List<ElementType> retList = new ArrayList<ElementType>();
		Finder finder = new Finder(Text.class);
		new TraversalUtil(documentPart.getContent(), finder);
		for (Object obj : finder.results) {
			Text text = (Text) obj;
			String name="";
			R parent1 = (R)text.getParent();
			if(parent1.getParent() instanceof P){
				P parent2 = (P)parent1.getParent();
				if(parent2.getParent() instanceof SdtContentBlock){
					SdtContentBlock parent3 = (SdtContentBlock)parent2.getParent();
					if(parent3.getParent() instanceof SdtBlock){
						SdtBlock parent4 = (SdtBlock)parent3.getParent();
						try{
							name = getBoxName(parent4.getSdtPr());
							if (showSysOut) System.out.println("getTextList - Adding TextElement: " + name + " / " + text.getValue());
							ElementType t = new ElementType(name, false,text.getValue(),1);
							retList.add(t);
						}catch(Exception e){
							
						}						
					}
				}
			}
		}
		finder = null;
		
		if (showSysOut) System.out.println("getTextList - retList: " + retList.size());
		return retList;
	}
	
	public List<ElementType> getTextListNoBookmark(){
		List<ElementType> retList = new ArrayList<ElementType>();
		Finder finder = new Finder(Text.class);
		new TraversalUtil(documentPart.getContent(), finder);
		for (Object obj : finder.results) {
			Text text = (Text) obj;			
			String name="";
			R parent1 = (R)text.getParent();
			if(parent1.getParent() instanceof SdtRun){
				SdtRun parent2 = (SdtRun)parent1.getParent();
				name = getBoxName(parent2.getSdtPr());
				if (showSysOut) System.out.println("getTextListNoBookmark - Adding TextElement: " + name + " / " + text.getValue());
				ElementType t = new ElementType(name, false,text.getValue(),1);
				retList.add(t);
			}
		}
		finder = null;
		
		if (showSysOut) System.out.println("getTextListNoBookmark - retList: " + retList.size());
		return retList;
	}
	
	
	
	public boolean isChecked(List<ElementType> sourceList, String param) {
		if ("".equals(param)) return false;

		boolean retVal = false;
		for (ElementType el : sourceList) {
			if (param.equalsIgnoreCase(el.getName())) {
				retVal = el.isChecked();
				break;
			}
		}
		return retVal;
	}
	
	public String getTextValue(List<ElementType> sourceList, String param) {
		if ("".equals(param)) return "";
		
		String retVal = "";
		for (ElementType el : sourceList) {
			if (param.equalsIgnoreCase(el.getName())) {
				retVal += (retVal.equals("")) ? el.getText() : ", " + el.getText();
			}
		}
		return retVal;
	}
	
}