package biz.webgate.documenthandler.test;

import java.util.List;

import org.apache.log4j.Level;
import org.apache.log4j.Logger;
import org.docx4j.Docx4jProperties;
import org.junit.Test;

import biz.webgate.documenthandler.ElementType;
import biz.webgate.documenthandler.WordHandler;
public class TestDocM {
	@Test
	public void testActiveXElement(){
		// disable DocX4j logging
		Logger.getRootLogger().setLevel(Level.ERROR);
		Docx4jProperties.getProperties().setProperty("docx4j.Log4j.Configurator.disabled", "true");
				
				
		//String fileName = "C:\\SVGer\\F_Import\\YV201500002_20160607_IMP_URTEIL.docm";
		String fileName = "C:\\SVGer\\F_Import\\mb199000100_20161206_imp_urteil.docm";
		
		
		WordHandler handler = new WordHandler();
		handler.read(fileName);
		
		List<ElementType> cbList = handler.getCBs();
		
		List<ElementType> tList = handler.getTextList();
		tList.addAll(handler.getTextListNoBookmark());
		
		System.out.println("CBoxes: " + cbList.size() + "  / TList: " + tList.size());
		
		/*
		String g_nr = handler.getTextValue(tList, "G_NR");
		System.out.println("G_NR: " + g_nr);
		String ahv = handler.getTextValue(tList, "AHV_NR");
		System.out.println("AHV_NR: " + ahv);
		String g_nr_vi = handler.getTextValue(tList, "G_NR_VI");
		System.out.println("G_NR_VI: " + g_nr_vi);
		String kammer = handler.getTextValue(tList, "KAMMER");
		System.out.println("KAMMER: " + kammer);
		String teilnehmende = handler.getTextValue(tList, "TEILNEHMENDE");
		System.out.println("TEILNEHMENDE: " + teilnehmende);
		String entscheid = handler.getTextValue(tList, "ENTSCHEID");
		System.out.println("ENTSCHEID: " + entscheid);
		String entscheiddatum = handler.getTextValue(tList, "ENTSCHEIDDATUM");
		System.out.println("ENTSCHEIDDATUM: " + entscheiddatum);
		
		String kbeschrieb = handler.getTextValue(tList, "boxKurzbeschrieb");
		System.out.println("boxKurzbeschrieb: " + kbeschrieb);
		
		boolean isChecked = handler.isChecked(cbList, "boxExtern");
		System.out.println("boxExtern : " + isChecked);
		isChecked = handler.isChecked(cbList, "boxIntern");
		System.out.println("boxIntern : " + isChecked);
		isChecked = handler.isChecked(cbList, "boxKeine");
		System.out.println("boxKeine : " + isChecked);
		isChecked = handler.isChecked(cbList, "boxAnwendungsfall");
		System.out.println("boxAnwendungsfall : " + isChecked);
		isChecked = handler.isChecked(cbList, "boxHinweisfall");
		System.out.println("boxHinweisfall : " + isChecked);
		isChecked = handler.isChecked(cbList, "boxZwischenentscheid");
		System.out.println("boxZwischenentscheid : " + isChecked);
		
		isChecked = handler.isChecked(cbList, "boxAbweisung");
		System.out.println("boxAbweisung : " + isChecked);
		isChecked = handler.isChecked(cbList, "boxTeilGut");
		System.out.println("boxTeilGut : " + isChecked);
		isChecked = handler.isChecked(cbList, "boxGutRück");
		System.out.println("boxGutRück : " + isChecked);
		isChecked = handler.isChecked(cbList, "boxGutheissung");
		System.out.println("boxGutheissung : " + isChecked);
		isChecked = handler.isChecked(cbList, "boxAndere");
		System.out.println("boxAndere : " + isChecked);
		*/
	}
}
