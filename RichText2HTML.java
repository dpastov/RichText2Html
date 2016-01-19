/*
 *  lnrt2html - Lotus Notes Rich Text to HTML Converter
 * 
 *  Copyright (c) 2011 Tran Dinh Thoai <dthoai@yahoo.com>
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU General Public License
 * version 3.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this library; if not, write to the Free Software
 * Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
 */

package sirius.utils.domino;

import java.io.ByteArrayOutputStream;
import java.io.StringReader;
import java.util.HashMap;
import java.util.Map;
import java.util.Random;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import lotus.domino.DxlExporter;
import lotus.domino.Item;
import lotus.domino.RichTextItem;

import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;

public class RichText2Html {

	public static final int ALL_OPTIONS_OFF = 0;
	public static final int USE_INLINE_STYLES = 1;
	
	private static final String ATTRIBUTE_STYLE = "style";
	
	private static final String ELEMENT_BOLD = "strong";
	private static final String ELEMENT_ITALIC = "i";
	private static final String ELEMENT_UNDERLINE = "u";
	private static final String ELEMENT_STRIKETHROUGH = "del";
	private static final String ELEMENT_SUPERSCRIPT = "sup";
	private static final String ELEMENT_SUBSCRIPT = "sub";
	
	private String richText;
	private String plainText;
	private final int options;

	private Map<String, ParDef> parDefs;
	 
	private RichText2Html(String richText, String plainText) {
		this.richText = richText;
		this.plainText = plainText;
		this.options = ALL_OPTIONS_OFF;
	}
	
	public RichText2Html(Item item) {
		this(item, ALL_OPTIONS_OFF);
	}
	
	public RichText2Html(Item item, int options) {
		this.options = options;

		try {
			this.plainText = "";
			
			if (item.getType() == 1) {
				RichTextItem rtitem = (RichTextItem)item;
				this.plainText = rtitem.getUnformattedText();
			}
			else {
				this.plainText = item.getValueString();
			}

			DxlExporter dxl = item.getParent().getParentDatabase().getParent()
					.createDxlExporter();
			dxl.setConvertNotesBitmapsToGIF(true);
			String html = dxl.exportDxl(item.getParent());
			dxl.recycle();
			int pos = html.indexOf("<item");
			if (pos < 0) {
				html = "";
			} else {
				html = html.substring(pos);
				html = html.substring(0, html.lastIndexOf("</document>"));
			}

			this.richText = this.plainText;
			Document dom = loadDOM(html);
			NodeList nodes = dom.getFirstChild().getChildNodes();
			for (int j = 0; j < nodes.getLength(); j++) {
				Node node = nodes.item(j);
				if (node.getNodeName() == "item") {
					Node attrNode = node.getAttributes().getNamedItem("name");
					if (attrNode != null) {
						if (attrNode.getNodeValue().equals(item.getName())) {
							richText = saveDOM(node);
							String pattern = "<richtext>";
							pos = richText.indexOf(pattern);
							richText = richText.substring(pos
									+ pattern.length());
							pos = richText.lastIndexOf("</richtext>");
							richText = richText.substring(0, pos);
							break;
						}
					}
				}
			}

		} catch (Exception e) {
			this.richText = this.plainText;
		}
	}

	public String parse() {
		String html = plainText;

		try {
			Document doc = loadDOM(richText);
			loadParDefs(doc);
			transform(doc.getDocumentElement());
			html = saveDOM(doc);
		} catch (Exception e) {
			html = plainText;
		}

		return html;
	}

	private void loadParDefs(Document doc) throws Exception {
		parDefs = new HashMap<String, ParDef>();
		NodeList nodes = doc.getElementsByTagName("pardef");
		for (int i = nodes.getLength() - 1; i >= 0; i--) {
			Node node = nodes.item(i);
			ParDef def = new ParDef();

			NamedNodeMap attrs = node.getAttributes();
			Node attr = attrs.getNamedItem("id");
			if (attr != null) {
				String id = attr.getNodeValue();
				boolean found = false;

				if (!found) {
					attr = attrs.getNamedItem("list");
					if (attr != null) {
						found = true;
						def.kind = ParDef.LIST;
						def.style = attr.getNodeValue();
						attr = attrs.getNamedItem("leftmargin");
						if (attr != null) {
							def.leftMargin = attr.getNodeValue();
						}
						attr = attrs.getNamedItem("align");
						if (attr != null) {
							def.align = attr.getNodeValue();
						}
					}
				}

				if (!found) {
					attr = attrs.getNamedItem("align");
					if (attr != null) {
						found = true;
						def.kind = ParDef.PARAGRAPH;
						def.align = attr.getNodeValue();
					}
					attr = attrs.getNamedItem("leftmargin");
					if (attr != null) {
						def.leftMargin = attr.getNodeValue();
					}
					attr = attrs.getNamedItem("spaceafter");
					if (attr != null) {
						def.spaceAfter = attr.getNodeValue();
					}
				}

				if (!found) {
					attr = attrs.getNamedItem("newpage");
					if (attr != null) {
						def.newPage = attr.getNodeValue();
					}
				}

				parDefs.put(id, def);

				node.getParentNode().removeChild(node);
			}
		}

	}

	private void transform(Node parent) throws Exception {
		NodeList children;
		Node curNode;
		NamedNodeMap attrs;
		Node attr;
		ParDef def;
		Node newNode = null;
		Node child;
		NodeList curChildren;
		String id;
		String prvId = "";
		Node prvNode = null;

		children = parent.getChildNodes();
		for (int i = 0; i < children.getLength(); i++) {
			curNode = children.item(i);
			if (curNode.getNodeName() == "par") {
				attrs = curNode.getAttributes();
				attr = attrs.getNamedItem("def");
				id = attr.getNodeValue();
				def = parDefs.get(id);
				if (def.kind == ParDef.LIST) {
					if (prvId.equals(id)) {
						newNode = prvNode;
					} else {
						newNode = createList(curNode, def);
						prvId = id;
						prvNode = newNode;
					}
					child = parent.getOwnerDocument().createElement("li");
					newNode.appendChild(child);
					
					curChildren = curNode.getChildNodes();
					int curChildrenLen = curChildren.getLength();
					for (int j = 0; j < curChildrenLen; j++) {
						child.appendChild(curChildren.item(j).cloneNode(true));
					}
				}
				if (def.kind == ParDef.PARAGRAPH) {
					newNode = createParagraph(curNode, def);
					curChildren = curNode.getChildNodes();
					int curChildrenLen = curChildren.getLength();
					for (int j = 0; j < curChildrenLen; j++) {
						newNode.appendChild(curChildren.item(j).cloneNode(true));
					}
				}
				parent.replaceChild(newNode, curNode);
			} else if (curNode.getNodeName() == "break") {
				newNode = parent.getOwnerDocument().createElement("br");
				parent.replaceChild(newNode, curNode);
			} else if (curNode.getNodeName() == "run") {
				newNode = createRun(curNode);
				if (newNode == null) {
					parent.removeChild(curNode);
				} else {
					parent.replaceChild(newNode, curNode);
				}
			} else if (curNode.getNodeName() == "horizrule") {
				newNode = createRule(curNode);
				parent.replaceChild(newNode, curNode);
			} else if (curNode.getNodeName() == "section") {
				newNode = createSection(curNode);
				parent.replaceChild(newNode, curNode);
			} else if (curNode.getNodeName() == "computedtext") {
				parent.removeChild(curNode);
			} else if (curNode.getNodeName() == "urllink") {
				newNode = createUrlLink(curNode);
				parent.replaceChild(newNode, curNode);
			} else if (curNode.getNodeName() == "popup") {
				newNode = createPopup(curNode);
				parent.replaceChild(newNode, curNode);
			} else if (curNode.getNodeName() == "button") {
				newNode = createButton(curNode);
				parent.replaceChild(newNode, curNode);
			} else if (curNode.getNodeName() == "actionhotspot") {
				newNode = createActionHotspot(curNode);
				parent.replaceChild(newNode, curNode);
			} else if (curNode.getNodeName() == "table") {
				boolean found = false;
				Node center = curNode.getParentNode();
				if (center != null) {
					if (center.getNodeName() == "center") {
						Node attrNode = center.getAttributes().getNamedItem(
								"table");
						if (attrNode != null) {
							if ("true".equals(attrNode.getNodeValue())) {
								found = true;
							}
						}
					}
				}
				if (!found) {
					newNode = createTable(curNode);
					parent.replaceChild(newNode, curNode);
				}
			} else if (curNode.getNodeName() == "picture") {
				parent.removeChild(curNode);
			}
		}

		children = parent.getChildNodes();
		for (int i = 0; i < children.getLength(); i++) {
			curNode = children.item(i);
			transform(curNode);
		}
	}

	private Node createTable(Node node) {
		Node center = null;
		Node tag = node.getOwnerDocument().createElement("table");

		Map<String, String> properties = new HashMap<String, String>();

		Attr attr = node.getOwnerDocument().createAttribute("cellspacing");
		attr.setNodeValue("0");
		tag.getAttributes().setNamedItem(attr);
		attr = node.getOwnerDocument().createAttribute("cellpadding");
		attr.setNodeValue("0");
		tag.getAttributes().setNamedItem(attr);

		String style = "";
		String borderColor = "black";
		Node attrNode = node.getAttributes().getNamedItem("cellbordercolor");
		if (attrNode != null) {
			borderColor = attrNode.getNodeValue();
		}
		properties.put("BorderColor", borderColor);
		String borderStyle = "solid";
		attrNode = node.getAttributes().getNamedItem("cellborderstyle");
		if (attrNode != null) {
			borderStyle = attrNode.getNodeValue();
		}
		properties.put("BorderStyle", borderStyle);
		String widthType = "";
		String tableWidth = "";
		String refWidth = "";
		attrNode = node.getAttributes().getNamedItem("widthtype");
		if (attrNode != null) {
			widthType = attrNode.getNodeValue();
		}
		attrNode = node.getAttributes().getNamedItem("refwidth");
		if (attrNode != null) {
			refWidth = attrNode.getNodeValue();
		}
		if (widthType.equals("fitmargins")) {
			tableWidth = "100%";
		} else if (widthType.equals("fitwindow")) {
			tableWidth = "100%";
		} else if (widthType.equals("fixedleft")) {
			tableWidth = refWidth;
			style += "float:left;";
		} else if (widthType.equals("fixedright")) {
			tableWidth = refWidth;
			style += "float:right;";
		} else if (widthType.equals("fixedcenter")) {
			tableWidth = refWidth;
			center = node.getOwnerDocument().createElement("center");
			attr = node.getOwnerDocument().createAttribute("table");
			attr.setNodeValue("true");
			center.getAttributes().setNamedItem(attr);
			center.appendChild(tag);
		}
		if (tableWidth.length() > 0) {
			style += "width:" + tableWidth + ";";
		}
		attrNode = node.getAttributes().getNamedItem("leftmargin");
		if (attrNode != null) {
			double margin = 0;
			try {
				margin = Double.parseDouble(attrNode.getNodeValue().replace(
						"in", ""));
				margin -= 1;
				if (margin < 0)
					margin = 0;
			} catch (Exception e) {
				margin = 0;
			}
			if (widthType.equals("fitmargins")) {
				style += "margin-left:" + margin + "in";
			}
		}

		if (isUseInlineStyles() && style.length() > 0) {
			attr = node.getOwnerDocument().createAttribute(ATTRIBUTE_STYLE);
			attr.setNodeValue(style);
			tag.getAttributes().setNamedItem(attr);
		}

		NodeList children = node.getChildNodes();
		int rowCount = 0;
		int colNo = 0;
		for (int i = 0; i < children.getLength(); i++) {
			Node child = children.item(i);
			if (child.getNodeName() == "tablecolumn") {
				colNo++;
				String key = "W" + colNo;
				String val = "";
				attrNode = child.getAttributes().getNamedItem("width");
				if (attrNode != null) {
					val = attrNode.getNodeValue();
				}
				properties.put(key, val);
			}
			if (child.getNodeName() == "tablerow") {
				rowCount++;
			}
		}
		properties.put("RowCount", rowCount + "");
		properties.put("CellCount", colNo + "");
		int rowNo = 0;
		for (int i = 0; i < children.getLength(); i++) {
			Node child = children.item(i);
			if (child.getNodeName() == "tablerow") {
				rowNo++;
				if (rowNo >= rowCount) {
					properties.put("LastRow", "true");
				} else {
					properties.put("LastRow", "false");
				}
				properties.put("RowNo", rowNo + "");
				tag.appendChild(createTableRow(child, properties));
			}
		}

		if (center == null) {
			return tag;
		} else {
			return center;
		}
	}

	private Node createTableRow(Node node, Map<String, String> properties) {
		Node tag = node.getOwnerDocument().createElement("tr");

		NodeList children = node.getChildNodes();
		int cellCount = 0;
		try {
			cellCount = Integer.parseInt(properties.get("CellCount"));
		} catch (Exception e) {
			cellCount = 0;
		}
		int rowNo = 0;
		try {
			rowNo = Integer.parseInt(properties.get("RowNo"));
		} catch (Exception e) {
			rowNo = 0;
		}
		int cellNo = 0;
		for (int i = 0; i < children.getLength(); i++) {
			Node child = children.item(i);
			if (child.getNodeName() == "tablecell") {
				cellNo++;
				String span = properties.get("span." + rowNo + "." + cellNo);
				while (span != null && "true".equals(span)) {
					cellNo++;
					span = properties.get("span." + rowNo + "." + cellNo);
				}
				if (cellNo >= cellCount) {
					properties.put("LastCell", "true");
				} else {
					properties.put("LastCell", "false");
				}
				properties.put("CellNo", cellNo + "");
				tag.appendChild(createTableCell(child, properties));
			}
		}

		return tag;
	}

	private Node createTableCell(Node node, Map<String, String> properties) {
		Node tag = node.getOwnerDocument().createElement("td");

		Attr attr;
		String style = "";
		String borderWidth = "";
		String borderColor = properties.get("BorderColor");
		String borderStyle = properties.get("BorderStyle");
		String cellWidth = properties.get("W" + properties.get("CellNo"));
		Node attrNode;

		int rowCount = 0;
		try {
			rowCount = Integer.parseInt(properties.get("RowCount"));
		} catch (Exception e) {
			rowCount = 0;
		}

		int colCount = 0;
		try {
			colCount = Integer.parseInt(properties.get("CellCount"));
		} catch (Exception e) {
			colCount = 0;
		}

		int rowNo = 0;
		try {
			rowNo = Integer.parseInt(properties.get("RowNo"));
		} catch (Exception e) {
			rowNo = 0;
		}
		int colNo = 0;
		try {
			colNo = Integer.parseInt(properties.get("CellNo"));
		} catch (Exception e) {
			colNo = 0;
		}

		attrNode = node.getAttributes().getNamedItem("rowspan");
		if (attrNode != null) {
			attr = node.getOwnerDocument().createAttribute("rowspan");
			attr.setNodeValue(attrNode.getNodeValue());
			tag.getAttributes().setNamedItem(attr);
			int rowspan = 0;
			try {
				rowspan = Integer.parseInt(attrNode.getNodeValue());
			} catch (Exception e) {
				rowspan = 0;
			}
			for (int i = 1; i < rowspan; i++) {
				properties.put("span." + (rowNo + i) + "." + colNo, "true");
			}
			if (rowNo + rowspan > rowCount) {
				properties.put("LastRow", "true");
			} else {
				properties.put("LastRow", "false");
			}
		}
		attrNode = node.getAttributes().getNamedItem("columnspan");
		if (attrNode != null) {
			attr = node.getOwnerDocument().createAttribute("colspan");
			attr.setNodeValue(attrNode.getNodeValue());
			tag.getAttributes().setNamedItem(attr);
			int colspan = 0;
			try {
				colspan = Integer.parseInt(attrNode.getNodeValue());
			} catch (Exception e) {
				colspan = 0;
			}
			for (int i = 1; i < colspan; i++) {
				properties.put("span." + rowNo + "." + (colNo + i), "true");
			}
			if (colNo + colspan > colCount) {
				properties.put("LastCell", "true");
			} else {
				properties.put("LastCell", "false");
			}
		}

		attrNode = node.getAttributes().getNamedItem("borderwidth");
		if (attrNode != null) {
			borderWidth = attrNode.getNodeValue();
		}
		if (borderWidth.length() > 0) {
			String[] fields = borderWidth.split(" ");
			if (fields.length > 0) {
				style += "border-top:" + borderStyle + " " + fields[0] + " "
						+ borderColor + ";";
			} else {
				style += "border-top:" + borderStyle + " 1px " + borderColor
						+ ";";
			}
			if (fields.length > 1) {
				if ("true".equals(properties.get("LastCell"))) {
					style += "border-right:" + borderStyle + " " + fields[1]
							+ " " + borderColor + ";";
				}
			} else {
				style += "border-right:" + borderStyle + " 1px " + borderColor
						+ ";";
			}
			if (fields.length > 2) {
				if ("true".equals(properties.get("LastRow"))) {
					style += "border-bottom:" + borderStyle + " " + fields[2]
							+ " " + borderColor + ";";
				}
			} else {
				style += "border-bottom:" + borderStyle + " 1px " + borderColor
						+ ";";
			}
			if (fields.length > 3) {
				style += "border-left:" + borderStyle + " " + fields[3] + " "
						+ borderColor + ";";
			} else {
				style += "border-left:" + borderStyle + " 1px " + borderColor
						+ ";";
			}
		} else {
			style += "border-top:" + borderStyle + " 1px " + borderColor + ";";
			style += "border-left:" + borderStyle + " 1px " + borderColor + ";";
			if ("true".equals(properties.get("LastRow"))) {
				style += "border-bottom:" + borderStyle + " 1px " + borderColor
						+ ";";
			}
			if ("true".equals(properties.get("LastCell"))) {
				style += "border-right:" + borderStyle + " 1px " + borderColor
						+ ";";
			}
		}

		if (cellWidth != null && cellWidth.length() > 0) {
			style += "width:" + cellWidth + ";";
		}

		if (isUseInlineStyles() && style.length() > 0) {
			attr = node.getOwnerDocument().createAttribute(ATTRIBUTE_STYLE);
			attr.setNodeValue(style);
			tag.getAttributes().setNamedItem(attr);
		}

		NodeList children = node.getChildNodes();
		for (int i = 0; i < children.getLength(); i++) {
			Node child = children.item(i);
			tag.appendChild(child.cloneNode(true));
		}

		return tag;
	}

	private Node createActionHotspot(Node node) {
		Node tag = node.getOwnerDocument().createElement("span");

		String style = "";
		String hotspotstyle = "";

		Node attrNode = node.getAttributes().getNamedItem("hotspotstyle");
		if (attrNode != null) {
			hotspotstyle = attrNode.getNodeValue();
		}

		if (!hotspotstyle.equals("none")) {
			style += "border:solid 1px teal;";
		}

		if (isUseInlineStyles() && style.length() > 0) {
			Attr attr = node.getOwnerDocument().createAttribute(ATTRIBUTE_STYLE);
			attr.setNodeValue(style);
			tag.getAttributes().setNamedItem(attr);
		}

		NodeList children = node.getChildNodes();
		for (int i = 0; i < children.getLength(); i++) {
			tag.appendChild(children.item(i).cloneNode(true));
		}

		return tag;
	}

	private Node createButton(Node node) {
		Node tag;
		Attr attr;
		Node attrNode;
		String text = "";
		String width = "";
		String widthType = "";
		String style = "";

		tag = node.getOwnerDocument().createElement("input");

		attr = node.getOwnerDocument().createAttribute("type");
		attr.setNodeValue("button");
		tag.getAttributes().setNamedItem(attr);

		NodeList children = node.getChildNodes();
		for (int i = 0; i < children.getLength(); i++) {
			Node child = children.item(i);
			if (child.getNodeName() == "code")
				continue;
			if (child.getNodeValue() != null) {
				text += child.getNodeValue();
			}
		}

		attr = node.getOwnerDocument().createAttribute("value");
		attr.setNodeValue(text);
		tag.getAttributes().setNamedItem(attr);

		attrNode = node.getAttributes().getNamedItem("width");
		if (attrNode != null) {
			width = attrNode.getNodeValue();
		}
		attrNode = node.getAttributes().getNamedItem("widthtype");
		if (attrNode != null) {
			widthType = attrNode.getNodeValue();
		}
		if (widthType.equals("fitcontent")) {
			width = "auto";
		}

		if (width.length() > 0) {
			style += "width:" + width + ";";
		}

		if (isUseInlineStyles() && style.length() > 0) {
			attr = node.getOwnerDocument().createAttribute(ATTRIBUTE_STYLE);
			attr.setNodeValue(style);
			tag.getAttributes().setNamedItem(attr);
		}

		return tag;
	}

	private Node createPopup(Node node) {
		Node tag = node.getOwnerDocument().createElement("span");

		Random random = new Random();
		String popupId = Long.toString(Math.abs(random.nextLong()), 36);

		Attr attr;
		String style = "cursor:pointer;cursor:hand;";

		Node title = node.getOwnerDocument().createElement("span");
		tag.appendChild(title);

		Node cover = node.getOwnerDocument().createElement("span");
		tag.appendChild(cover);
		attr = node.getOwnerDocument().createAttribute("id");
		attr.setNodeValue(popupId);
		cover.getAttributes().setNamedItem(attr);
		attr = node.getOwnerDocument().createAttribute("style");
		attr
				.setNodeValue("display:none;margin: 10px;border:solid 1px teal;width:300px;height:50px;");
		cover.getAttributes().setNamedItem(attr);

		NodeList children = node.getChildNodes();
		for (int i = 0; i < children.getLength(); i++) {
			Node child = children.item(i);
			if (child.getNodeName() == "popuptext") {
				NodeList curChildren = child.getChildNodes();
				for (int j = 0; j < curChildren.getLength(); j++) {
					cover.appendChild(curChildren.item(j).cloneNode(true));
				}
			} else if (child.getNodeName() != "code") {
				title.appendChild(child.cloneNode(true));
			}
		}

		NamedNodeMap attrs = node.getAttributes();
		Node attrNode = attrs.getNamedItem("hotspotstyle");
		if (attrNode != null) {
			String kind = attrNode.getNodeValue();
			if (kind.equals("highlight")) {
				style += "background-color:Lemonchiffon;";
			} else if (kind.equals("none")) {
			} else {
				style += "border:solid 1px teal;";
			}
		} else {
			style += "border:solid 1px teal;";
		}
		if (isUseInlineStyles()) {
			attr = node.getOwnerDocument().createAttribute(ATTRIBUTE_STYLE);
			attr.setNodeValue(style);
			title.getAttributes().setNamedItem(attr);
		}
		
		attrNode = attrs.getNamedItem("show");
		String event = "onmouseover";
		if (attrNode != null) {
			String kind = attrNode.getNodeValue();
			if (kind.equals("onclick")) {
				event = "onclick";
			}
		}
		attr = node.getOwnerDocument().createAttribute(event);
		attr.setNodeValue("var cover=document.getElementById('" + popupId
				+ "');cover.style.display='block';");
		title.getAttributes().setNamedItem(attr);
		attr = node.getOwnerDocument().createAttribute("onmouseout");
		attr.setNodeValue("var cover=document.getElementById('" + popupId
				+ "');cover.style.display='none';");
		title.getAttributes().setNamedItem(attr);

		return tag;
	}

	private Node createUrlLink(Node node) {
		Node tag = node.getOwnerDocument().createElement("a");

		Attr attr;
		String style = "";
		NamedNodeMap attrs = node.getAttributes();
		Node attrNode = attrs.getNamedItem("href");
		if (attrNode != null) {
			attr = node.getOwnerDocument().createAttribute("href");
			attr.setNodeValue(attrNode.getNodeValue());
			tag.getAttributes().setNamedItem(attr);
		}
		
		attrNode = attrs.getNamedItem("targetframe");
		if (attrNode != null) {
			attr = node.getOwnerDocument().createAttribute("target");
			attr.setNodeValue(attrNode.getNodeValue());
			tag.getAttributes().setNamedItem(attr);
		}
		
		attrNode = attrs.getNamedItem("showborder");
		if (attrNode != null) {
			if (attrNode.getNodeValue().equals("true")) {
				style += "border:solid 1px teal;";
			}
		}
		
		if (isUseInlineStyles() && style.length() > 0) {
			attr = node.getOwnerDocument().createAttribute(ATTRIBUTE_STYLE);
			attr.setNodeValue(style);
			tag.getAttributes().setNamedItem(attr);
		}
		
		NodeList children = node.getChildNodes();
		for (int i = 0; i < children.getLength(); i++) {
			Node child = children.item(i);
			if (child.getNodeName() == "code"){
				NodeList codeChildren = child.getChildNodes();
				for (int j = 0; j < codeChildren.getLength(); j++) {
					Node codeChild = codeChildren.item(j);
					if (codeChild.getNodeName() == "formula"){
						Node formulaNode = codeChild.getFirstChild();
						if (null != formulaNode) {
							String formula = formulaNode.getTextContent().toLowerCase();
							attr = node.getOwnerDocument().createAttribute("href");
							
							String res = formula.replace("\"", "");
							attr.setNodeValue(res);
							
							tag.getAttributes().setNamedItem(attr);
						}
					}
				}
				continue;
			}
			tag.appendChild(child.cloneNode(true));
		}

		return tag;
	}

	private Node createSection(Node node) {
		Node tag = node.getOwnerDocument().createElement("div");

		Random random = new Random();
		String coverId = Long.toString(Math.abs(random.nextLong()), 36);

		Node secTitle = node.getFirstChild().getNextSibling().getFirstChild()
				.getNextSibling().getFirstChild();
		Node title = node.getOwnerDocument().createElement("div");
		Node content = node.getOwnerDocument().createTextNode(
				secTitle.getNodeValue());
		title.appendChild(content);
		Attr attr = node.getOwnerDocument().createAttribute(ATTRIBUTE_STYLE);
		attr
				.setNodeValue("padding-left:5px;border-left:solid 5px teal;border-bottom:solid 1px teal;cursor:pointer;cursor:hand;");
		title.getAttributes().setNamedItem(attr);
		attr = node.getOwnerDocument().createAttribute("onclick");
		attr
				.setNodeValue("var sec=document.getElementById('"
						+ coverId
						+ "');if(sec.style.display=='none'){sec.style.display='block'}else{sec.style.display='none'}");
		title.getAttributes().setNamedItem(attr);
		tag.appendChild(title);

		Node cover = node.getOwnerDocument().createElement("div");
		attr = node.getOwnerDocument().createAttribute("id");
		attr.setNodeValue(coverId);
		cover.getAttributes().setNamedItem(attr);
		attr = node.getOwnerDocument().createAttribute(ATTRIBUTE_STYLE);
		attr.setNodeValue("display:none;");
		cover.getAttributes().setNamedItem(attr);
		tag.appendChild(cover);

		NodeList children = node.getChildNodes();
		for (int i = 2; i < children.getLength(); i++) {
			cover.appendChild(children.item(i).cloneNode(true));
		}
		cover.appendChild(node.getOwnerDocument().createElement("br"));

		return tag;
	}

	private Node createRule(Node node) {
		Node tag = node.getOwnerDocument().createElement("hr");
		Attr attr;
		Node attrNode;
		String style = "";

		attrNode = node.getAttributes().getNamedItem("height");
		if (attrNode != null) {
			style += "height:" + attrNode.getNodeValue() + ";";
		}
		attrNode = node.getAttributes().getNamedItem("width");
		if (attrNode != null) {
			style += "width:" + attrNode.getNodeValue() + ";";
		}
		attrNode = node.getAttributes().getNamedItem("color");
		if (attrNode != null) {
			style += "background-color:" + attrNode.getNodeValue() + ";";
		}

		if (isUseInlineStyles() && style.length() > 0) {
			attr = node.getOwnerDocument().createAttribute(ATTRIBUTE_STYLE);
			attr.setNodeValue(style);
			tag.getAttributes().setNamedItem(attr);
		}

		return tag;
	}

	private Node createRun(Node node) {
		Node tag = null;
		NamedNodeMap attrs;
		Node attrNode;
		NodeList children;
		Node newNode = null;
		Attr attr;
		String style = "";
		
		attrs = node.getAttributes();
		attrNode = attrs.getNamedItem("html");
		if (attrNode != null) {
			if (attrNode.getNodeValue().equals("true")) {
				newNode = node.getOwnerDocument().createElement("code");
				attrNode = attrs.getNamedItem("highlight");
				if (attrNode != null) {
					style += "background-color:"
							+ attrNode.getNodeValue()
									.replace("yellow", "Khaki").replace("blue",
											"Lightsteelblue").replace("pink",
											"Thistle") + ";";
				} else {
					style += "background-color:gainsboro;";
				}
			}
		}
		if (newNode == null) {
			attrNode = attrs.getNamedItem("highlight");
			if (attrNode != null) {
				newNode = node.getOwnerDocument().createElement("span");
				style += "background-color:"
						+ attrNode.getNodeValue().replace("yellow",
								"Lemonchiffon").replace("blue", "Lightcyan")
								.replace("pink", "Mistyrose") + ";";
			}
		}

		if (node.getChildNodes().getLength() < 2) {
			if (newNode == null) {
				newNode = node.getOwnerDocument().createElement("span");
			}
			if (isUseInlineStyles() && style.length() > 0) {
				attr = node.getOwnerDocument().createAttribute(ATTRIBUTE_STYLE);
				attr.setNodeValue(style);
				newNode.getAttributes().setNamedItem(attr);
			}
			children = node.getChildNodes();
			for (int i = 0; i < children.getLength(); i++) {
				newNode.appendChild(children.item(i).cloneNode(true));
			}
			tag = newNode;
		} else {
			Node font = node.getFirstChild();
			if (font.getNodeName() != "font") {
				if (newNode == null) {
					newNode = node.getOwnerDocument().createElement("span");
				}
				if (isUseInlineStyles() && style.length() > 0) {
					attr = node.getOwnerDocument().createAttribute(ATTRIBUTE_STYLE);
					attr.setNodeValue(style);
					newNode.getAttributes().setNamedItem(attr);
				}
				children = node.getChildNodes();
				for (int i = 0; i < children.getLength(); i++) {
					newNode.appendChild(children.item(i).cloneNode(true));
				}
				tag = newNode;
			} else {
				attrs = font.getAttributes();
				attrNode = attrs.getNamedItem(ATTRIBUTE_STYLE);
				Node inner = node.getOwnerDocument().createTextNode(node.getTextContent());
				Node styleNode = null;
				if (attrNode != null) {
					String fontStyle = attrNode.getNodeValue();
					
					if (fontStyle.indexOf("italic") >= 0) {
						styleNode = node.getOwnerDocument().createElement(ELEMENT_ITALIC);
						styleNode.appendChild(inner);
						inner = styleNode.cloneNode(true);
						
						style += "font-style:italic;";
					}
					
					if (fontStyle.indexOf("bold") >= 0) {
						styleNode = node.getOwnerDocument().createElement(ELEMENT_BOLD);
						styleNode.appendChild(inner);
						inner = styleNode.cloneNode(true);
						
						style += "font-weight:bold;";
					}
					
					if (fontStyle.indexOf("underline") >= 0) {
						styleNode = node.getOwnerDocument().createElement(ELEMENT_UNDERLINE);
						styleNode.appendChild(inner);
						inner = styleNode.cloneNode(true);
					}
					
					if (fontStyle.indexOf("strikethrough") >= 0) {
						styleNode = node.getOwnerDocument().createElement(ELEMENT_STRIKETHROUGH);
						styleNode.appendChild(inner);
						inner = styleNode.cloneNode(true);
					}
					
					if (fontStyle.indexOf("superscript") >= 0) {
						styleNode = node.getOwnerDocument().createElement(ELEMENT_SUPERSCRIPT);
						styleNode.appendChild(inner);
						inner = styleNode.cloneNode(true);
					}
					
					if (fontStyle.indexOf("subscript") >= 0) {
						styleNode = node.getOwnerDocument().createElement(ELEMENT_SUBSCRIPT);
						styleNode.appendChild(inner);
						inner = styleNode.cloneNode(true);
					}
					
					// TODO: Not implemented in html tags: emboss, shadow, extrude

					if (fontStyle.indexOf("emboss") >= 0) {
						if (null == styleNode) styleNode = inner.cloneNode(true);
					}
					
					if (fontStyle.indexOf("shadow") >= 0) {
						if (null == styleNode) styleNode = inner.cloneNode(true);
					}
					
					if (fontStyle.indexOf("extrude") >= 0) {
						if (null == styleNode) styleNode = inner.cloneNode(true);
					}
				}
				
				attrNode = attrs.getNamedItem("color");
				if (attrNode != null) {
					if (null == styleNode) styleNode = inner.cloneNode(true);
					style += "color:" + attrNode.getNodeValue() + ";";
				}
				attrNode = attrs.getNamedItem("size");
				if (attrNode != null) {
					if (null == styleNode) styleNode = inner.cloneNode(true);
					style += "font-size:" + attrNode.getNodeValue() + ";";
				}
				attrNode = attrs.getNamedItem("name");
				if (attrNode != null) {
					if (null == styleNode) styleNode = inner.cloneNode(true);
					style += "font-family:" + attrNode.getNodeValue() + ";";
				}

				if (null != styleNode)
					if (newNode == null) 
						newNode = styleNode;
					else
						newNode.appendChild(styleNode);
				
				if (isUseInlineStyles() && style.length() > 0) {
					attr = node.getOwnerDocument().createAttribute(ATTRIBUTE_STYLE);
					attr.setNodeValue(style);
					newNode.getAttributes().setNamedItem(attr);
				}
				
				children = node.getChildNodes();
				for (int i = 2; i < children.getLength(); i++) {
					newNode.appendChild(children.item(i).cloneNode(true));
				}
				tag = newNode;
			}
		}

		return tag;
	}

	private Node createParagraph(Node node, ParDef def) {
		Node tag = node.getOwnerDocument().createElement("p");
		String style = "";

		if (def.align.equals("full")) {
			style += "text-align: justify;";
		} else if (def.align.equals("center")) {
			style += "text-align: center;";
		} else if (def.align.equals("right")) {
			style += "text-align: right;";
		} else if (def.align.equals("none")) {
			style += "white-space:nowrap;";
		}

		String marginStyle = "";

		if (def.leftMargin.length() > 0 && node.getChildNodes().getLength() > 0) {
			double margin = 0;
			try {
				margin = Double
						.parseDouble(def.leftMargin.replaceAll("in", ""));
				margin -= 1;
			} catch (Exception e) {
				margin = 0;
			}
			marginStyle += "margin-left:" + margin + "in;";
		}

		if (def.spaceAfter.length() > 0 && node.getChildNodes().getLength() > 0) {
			if (def.spaceAfter.equals("2")) {
				marginStyle += "margin-bottom:1em";
			} else if (def.spaceAfter.equals("1.5")) {
				marginStyle += "margin-bottom:0.5em";
			}
		}

		if (marginStyle.length() > 0) {
			style += "margin:0px;" + marginStyle;
		}

		if (def.newPage.equals("true")) {
			style += "border-top: solid 1px black;";
		}

		if (isUseInlineStyles() && style.length() > 0) {
			Attr attr = node.getOwnerDocument().createAttribute(ATTRIBUTE_STYLE);
			attr.setNodeValue(style);
			tag.getAttributes().setNamedItem(attr);
		} else {
			if (node.getChildNodes().getLength() == 0) {
				tag = node.getOwnerDocument().createElement("br");
			}
		}

		return tag;
	}

	private Node createList(Node node, ParDef def) {
		Node tag = node.getOwnerDocument().createElement("ul");
		String style = "";

		if (def.style.equals("square")) {
			style += "list-style-type:square;";
		} else if (def.style.equals("circle")) {
			style += "list-style-type:circle;";
		} else if (def.style.equals("uncheck")) {
			style += "list-style-type: square;";
		} else if (def.style.equals("number")) {
			tag = node.getOwnerDocument().createElement("ol");
			style += "list-style-type:decimal;";
		} else if (def.style.equals("alphaupper")) {
			tag = node.getOwnerDocument().createElement("ol");
			style += "list-style-type:upper-alpha;";
		} else if (def.style.equals("alphalower")) {
			tag = node.getOwnerDocument().createElement("ol");
			style += "list-style-type:lower-alpha;";
		} else if (def.style.equals("romanupper")) {
			tag = node.getOwnerDocument().createElement("ol");
			style += "list-style-type:upper-roman;";
		} else if (def.style.equals("romanlower")) {
			tag = node.getOwnerDocument().createElement("ol");
			style += "list-style-type:lower-roman;";
		} else if (def.style.equals("bullet")) {
			style += "list-style-type:disc;";
		} else {
			style += "list-style-type:disc;";
		}

		double margin = 0;
		try {
			margin = Double.parseDouble(def.leftMargin.replaceAll("in", ""));
			margin -= 1.5;
		} catch (Exception e) {
			margin = 0;
		}
		style += "margin:0px;margin-left:" + margin + "in;";

		if (def.align.equals("full")) {
			style += "text-align: justify;";
		} else if (def.align.equals("center")) {
			style += "text-align: center;";
		} else if (def.align.equals("right")) {
			style += "text-align: right;";
		}

		if (isUseInlineStyles()) {
			Attr attr = node.getOwnerDocument().createAttribute(ATTRIBUTE_STYLE);
			attr.setNodeValue(style);
			tag.getAttributes().setNamedItem(attr);
		}
		
		return tag;
	}

	private String saveDOM(Document doc) throws Exception {
		String tag = saveDOM(doc.getDocumentElement());
		String pattern = "<body>";
		int pos = tag.indexOf(pattern);
		tag = tag.substring(pos + pattern.length());
		pos = tag.lastIndexOf("</body>");
		tag = tag.substring(0, pos);
		return tag;
	}

	private static String saveDOM(Node node) throws Exception {
		TransformerFactory factory = TransformerFactory.newInstance();
		Transformer transformer = factory.newTransformer();

		DOMSource source = new DOMSource(node);
		ByteArrayOutputStream stream = new ByteArrayOutputStream();
		StreamResult result = new StreamResult(stream);
		transformer.transform(source, result);

		String tag = stream.toString("UTF-8");

		int pos = tag.indexOf("<" + node.getNodeName());
		tag = tag.substring(pos);

		return tag;
	}

	private static Document loadDOM(String source) throws Exception {
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		dbf.setValidating(false);
		dbf.setNamespaceAware(false);
		dbf.setIgnoringComments(false);
		dbf.setIgnoringElementContentWhitespace(false);
		dbf.setExpandEntityReferences(false);
		DocumentBuilder db = dbf.newDocumentBuilder();
		return db.parse(new InputSource(new StringReader("<body>" + source
				+ "</body>")));
	}

	private boolean isOptionSet(int option) {
		return (this.options & option) == option;
	}
	
	private boolean isUseInlineStyles() {
		return isOptionSet(USE_INLINE_STYLES);
	}
	
	public static String parse(String richText, String plainText) {
		return new RichText2Html(richText, plainText).parse();
	}

	public static String parse(Item item) {
		return new RichText2Html(item).parse();
	}
	
	public static String parse(Item item, int options) {
		return new RichText2Html(item, options).parse();
	}

	public static class ParDef {

		public static final byte PARAGRAPH = 1;
		public static final byte LIST = 2;

		public byte kind = PARAGRAPH;
		public String style = "";
		public String leftMargin = "";
		public String align = "";
		public String spaceAfter = "";
		public String newPage = "";

	}

	public static String getDxl(Item item) {
		String richText = "";

		try {
			DxlExporter dxl = item.getParent().getParentDatabase().getParent().createDxlExporter();
			dxl.setConvertNotesBitmapsToGIF(true);
			String html = dxl.exportDxl(item.getParent());
			dxl.recycle();
			int pos = html.indexOf("<item");
			if (pos < 0) {
				html = "";
			} else {
				html = html.substring(pos);
				html = html.substring(0, html.lastIndexOf("</document>"));
			}

			Document dom = loadDOM(html);
			NodeList nodes = dom.getFirstChild().getChildNodes();
			for (int j = 0; j < nodes.getLength(); j++) {
				Node node = nodes.item(j);
				if (node.getNodeName() == "item") {
					Node attrNode = node.getAttributes().getNamedItem("name");
					if (attrNode != null) {
						if (attrNode.getNodeValue().equals(item.getName())) {
							richText = saveDOM(node);
							String pattern = "<richtext>";
							pos = richText.indexOf(pattern);
							richText = richText.substring(pos
									+ pattern.length());
							pos = richText.lastIndexOf("</richtext>");
							richText = richText.substring(0, pos);
							break;
						}
					}
				}
			}

		} catch (Exception e) {
		}

		return richText;
	}

}
