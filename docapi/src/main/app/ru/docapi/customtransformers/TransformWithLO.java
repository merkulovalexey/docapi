package ru.docapi.customtransformers;

import java.io.IOException;

import org.mule.api.MuleMessage;
import org.mule.api.transformer.TransformerException;
import org.mule.module.json.JsonData;
import org.mule.transformer.AbstractMessageTransformer;
import org.mule.util.Base64;


public class TransformWithLO extends AbstractMessageTransformer{

	@Override
	public Object transformMessage(MuleMessage message, String outputEncoding)
			throws TransformerException {
		// TODO Auto-generated method stub
		
		JsonData payload = (JsonData) message.getPayload(); 
		
		String base64string = payload.getAsString("template");
		byte[] byteArray = Base64.decode(base64string);
		
		LibreOfficeWorker libreOfficeWorker = new LibreOfficeWorker();
		byte[] resultByteArray = new byte[0];
		try {
			libreOfficeWorker.connect("uno:socket,host=localhost,port=2083;urp;StarOffice.ServiceManager");
			resultByteArray = libreOfficeWorker.fillTemplate(byteArray);
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		String resultBase64string;
		
		try {
			resultBase64string = Base64.encodeBytes(resultByteArray);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			resultBase64string = null;
		}
		
		message.setPayload(resultBase64string);
		System.out.println( resultBase64string );
		
		return null;
	}

}
