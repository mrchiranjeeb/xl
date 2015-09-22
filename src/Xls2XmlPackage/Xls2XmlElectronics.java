package Xls2XmlPackage;

import java.io.File;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import jxl.CellType;
import jxl.DateCell;
import jxl.Sheet;
import jxl.Workbook;

public class Xls2XmlElectronics{
	public void mainf(String inFileName, String outFolderName) throws Exception{
		File f=new File(inFileName);//creates file		
		System.out.println("excel file loaded");
		Workbook wb= null;
		try{
		 wb=Workbook.getWorkbook(f);
		System.out.println("workbook loaded");
		}catch(Exception e){
			System.out.println("Input file or path is invalid.. please check.." + e);
			return;
		}
		Sheet s=wb.getSheet(0);
		int rowNum=s.getRows();
		int colNum=s.getColumns();

		int i;
		for(i=6;i<rowNum;i++){
			String xmlOutput=outFolderName+"/"+(s.getCell(2, i).getContents()+"_"+s.getCell(3, i).getContents()+"_"+s.getCell(4, i).getContents())+".xml";
			DocumentBuilderFactory dbf=DocumentBuilderFactory.newInstance();
			DocumentBuilder db=dbf.newDocumentBuilder();
			Document document=db.newDocument();
			Element root=document.createElement("products");
			document.appendChild(root);
			Attr attr=document.createAttribute("xmlns:xsi");
			attr.setValue("http://www.w3.org/2001/XMLSchema-instance");
			root.setAttributeNode(attr);
			
			Attr attr1=document.createAttribute("xsi:noNamespaceSchemaLocation");
			attr1.setValue("products-v0.7.xsd");
			root.setAttributeNode(attr1);
			
			Element produc=document.createElement("product");
			root.appendChild(produc);
			

			for(int j=4;j<=11;j++){
				
				Element member1=document.createElement((""+s.getCell(j,3).getContents()).trim());
				member1.appendChild(document.createTextNode((""+s.getCell(j, i).getContents()).trim()));
				produc.appendChild(member1);
		
				
			
			}
			//----------------product features-----------------------
			
			Element productFeatures=document.createElement("productFeatures");
			produc.appendChild(productFeatures);
			for(int j=12;j<=255;j++){
				
				if(s.getCell(j,i).getContents()!=""){
				
				Element productFeature=document.createElement("productFeature");
				productFeatures.appendChild(productFeature);
				
				Element qualifier=document.createElement("qualifier");
				qualifier.appendChild(document.createTextNode((""+s.getCell(j,3).getContents()).trim()));
				productFeature.appendChild(qualifier);
				
				Element value=document.createElement("value");
				value.appendChild(document.createTextNode((""+s.getCell(j, i).getContents()).trim()));
				productFeature.appendChild(value);
				}
			}
//				====================================================================================
//				========================================Sheet2 starts===============================
//				====================================================================================
				
				s=wb.getSheet(1);
				for(int j=0;j<=35;j++){
					if(s.getCell(j,i).getContents()!=""){
						
						Element productFeature=document.createElement("productFeature");
						productFeatures.appendChild(productFeature);
						
						Element qualifier=document.createElement("qualifier");
						qualifier.appendChild(document.createTextNode((""+s.getCell(j,3).getContents()).trim()));
						productFeature.appendChild(qualifier);
						
						Element value=document.createElement("value");
						value.appendChild(document.createTextNode((""+s.getCell(j, i).getContents()).trim()));
						productFeature.appendChild(value);
				}
				
				
				}
				
				
				

			
			//----------------Global Identifier-----------------------
			Element global=document.createElement("globalIdentifier");
			produc.appendChild(global);
			
				
				
				
				Element qualifier=document.createElement("qualifier");
				qualifier.appendChild(document.createTextNode((""+s.getCell(36, i).getContents()).trim()));
				global.appendChild(qualifier);
				
				Element value=document.createElement("value");
				value.appendChild(document.createTextNode((""+s.getCell(37, i).getContents()).trim()));
				global.appendChild(value);
				
				//----------------Seller Info-----------------------
				int j=38;
				Element sellerInfo=document.createElement("sellerInfo");
				produc.appendChild(sellerInfo);
				for(;j<=44;j++){//j=128
					if(j==42 || j==43 || j==44){
					DateCell dCell=null;
					String dateStr="",prefixD="",prefixM="";
					if(s.getCell(j, i).getType() == CellType.DATE && s.getCell(j,i).getContents()!=null){
						dCell = (DateCell)s.getCell(j, i);
						String dateS = ""+dCell.getDate();
					
						prefixD=(dCell.getDate().getDate()<10)?"0"+dCell.getDate().getDate():""+dCell.getDate().getDate();
						prefixM= ((dCell.getDate().getMonth()+1)<10)?"0"+(dCell.getDate().getMonth()+1):""+(dCell.getDate().getMonth()+1);
					dateStr = dateS.substring(dateS.length()-4, dateS.length())+"-"+prefixM+"-"+prefixD;
					}
					Element member2=document.createElement((""+s.getCell(j,3).getContents()).trim());
					member2.appendChild(document.createTextNode(dateStr));
					sellerInfo.appendChild(member2);
					}else{
						//System.out.println("I a here with j="+j+" and i="+i);
					Element member2=document.createElement((""+s.getCell(j,3).getContents()).trim());
					member2.appendChild(document.createTextNode((""+s.getCell(j, i).getContents()).trim()));
					sellerInfo.appendChild(member2);
					}
					
				
				
				}
				
				//----------------variant-----------------------
				/*Element variant=document.createElement("variant");
				produc.appendChild(variant);
				for(;j<139;j++){//j=136
					
					Element member3=document.createElement((""+s.getCell(j,3).getContents()).trim());
					member3.appendChild(document.createTextNode((""+s.getCell(j, i).getContents()).trim()));
					variant.appendChild(member3);
					
				
				
				}*/
				//---------------- article mini description and review added to product-----------------------
			for (; j <= 46; j++) {// j=139
				if(s.getCell(j,i).getContents()!=""){
				Element member4 = document.createElement((""+s.getCell(j,3).getContents()).trim());
				member4.appendChild(document.createTextNode((""+s.getCell(j, i).getContents()).trim()));
				produc.appendChild(member4);
				}

			}
			
//----------------medias-----------------------
			
			Element medias=document.createElement("medias");
			produc.appendChild(medias);
			for(;j<=50;j++){//j=141
				
				Element media=document.createElement("media");
				medias.appendChild(media);

				Element type=document.createElement("type");
				type.appendChild(document.createTextNode("Image"));
				media.appendChild(type);
				
				Element code=document.createElement("code");
				code.appendChild(document.createTextNode((""+s.getCell(j, i).getContents()).trim()));
				media.appendChild(code);
				
				
			}
			
			//----------------brand-----------------------	
			Element brand=document.createElement("brand");
			produc.appendChild(brand);
			
			for(j=55;j<=57;j++){//j=149
				
				Element member5=document.createElement((""+s.getCell(j,3).getContents()).trim());
				member5.appendChild(document.createTextNode((""+s.getCell(j, i).getContents()).trim()));
				brand.appendChild(member5);
							
				
				
			}
			//----------------rich attributes-----------------------	
			Element richattr=document.createElement("richAttribute");
			produc.appendChild(richattr);
			
			for(;j<=74;j++){//j=151
				
				Element member6=document.createElement((""+s.getCell(j,3).getContents()).trim());
				member6.appendChild(document.createTextNode((""+s.getCell(j, i).getContents()).trim()));
				richattr.appendChild(member6);
							
				
				
			}
			//----------------cross sell info up sell info-----------------------	
						
			for(;j<=78;j++){//j=168
				if(s.getCell(j,i).getContents()!=""){
				Element member7=document.createElement((""+s.getCell(j,3).getContents()).trim());
				
				produc.appendChild(member7);
				Element reference=document.createElement("reference");
				reference.appendChild(document.createTextNode((""+s.getCell(j, i).getContents()).trim()));
				member7.appendChild(reference);
							
				}
				
			}
			
			//----------------seo contents-----------------------	
			Element seoContent=document.createElement("seoContent");
			produc.appendChild(seoContent);
			
			for(;j<=82;j++){//j=172
				if(s.getCell(j,i).getContents()!=""){
				Element member8=document.createElement((""+s.getCell(j,3).getContents()).trim());
				member8.appendChild(document.createTextNode((""+s.getCell(j, i).getContents()).trim()));
				seoContent.appendChild(member8);
				}	
				
				
			}
			//----------------message-----------------------	
			
			Element messages=document.createElement("messages");
			produc.appendChild(messages);
			
			for(;j<=84;j++){//j=176
				if(s.getCell(j,i).getContents()!=""){
				Element member9=document.createElement((""+s.getCell(j,3).getContents()).trim());
				member9.appendChild(document.createTextNode((""+s.getCell(j, i).getContents()).trim()));
				messages.appendChild(member9);
							
				}
			}
				
			
				
			//----------------------closing--------------------
			try{
			TransformerFactory tf=TransformerFactory.newInstance();
			Transformer t=tf.newTransformer();
			DOMSource domS=new DOMSource(document);
			StreamResult sr=new StreamResult(new File(xmlOutput));
			t.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");
			t.setOutputProperty(OutputKeys.INDENT, "yes");
			t.transform(domS, sr);
			}catch(Exception e){
				System.out.println("Output path is incorrect.. please check");
				return;
			}
			
			s=wb.getSheet(0);
	}
		
		System.out.println((i-6) + " files created successfully.");


	}

}
