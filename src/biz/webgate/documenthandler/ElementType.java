package biz.webgate.documenthandler;

public class ElementType {
	private String name;
	private boolean checked;
	private String text;
	private int type;
	
	public ElementType(String name,boolean checked,String text,int type){
		this.name = name;
		this.checked = checked;
		this.text = text;
		this.type = type;
	}
	public String getName() {
		return name;
	}
	public boolean isChecked() {
		return checked;
	}
	public String getText() {
		return text;
	}
	public int getType() {
		return type;
	}
	
}
