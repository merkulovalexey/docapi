package ru.docapi.customtransformers;


import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.Date;
import com.sun.star.util.XCloseable;
import com.sun.star.util.XReplaceDescriptor;
import com.sun.star.util.XReplaceable;

public class LibreOfficeWorker {
	private XComponentContext xComponentContext;
	
	public void connect(String unoUrl) throws java.lang.Exception {
		
		unoUrl = "uno:socket,host=localhost,port=2083;urp;StarOffice.ServiceManager";
	 // create default local component context
        XComponentContext xLocalContext =com.sun.star.comp.helper.Bootstrap.createInitialComponentContext(null);

       // initial serviceManager
        XMultiComponentFactory xLocalServiceManager = xLocalContext.getServiceManager();
		
    // create a urlresolver
        Object urlResolver  = xLocalServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", xLocalContext );

       // query for the XUnoUrlResolver interface
        XUnoUrlResolver xUrlResolver = UnoRuntime.queryInterface( XUnoUrlResolver.class, urlResolver );
       
    
       // Import the object
        Object rInitialObject = xUrlResolver.resolve( unoUrl );

        // XComponentContext
        if( null != rInitialObject )
        {
        	System.out.println( "initial object successfully retrieved" );
        	
        	XMultiComponentFactory xOfficeFactory = (XMultiComponentFactory) UnoRuntime.queryInterface(XMultiComponentFactory.class, rInitialObject);
    	    
			 // retrieve the component context as property (it is not yet exported from the office)
			  // Query for the XPropertySet interface.
			XPropertySet xProperySet = (XPropertySet) UnoRuntime.queryInterface(  XPropertySet.class, xOfficeFactory);
			
			// Get the default context from the office server.
			Object oDefaultContext = xProperySet.getPropertyValue("DefaultContext"); 
			  
			// Query for the interface XComponentContext.
			xComponentContext = (XComponentContext) UnoRuntime.queryInterface( XComponentContext.class, oDefaultContext);
        	
        }
        else
        {
            System.out.println( "given initial-object name unknown at server side" );

        }
        
        
        
        
	}
	
	public byte[] fillTemplate(byte[] byteArray) throws Exception {
		
		
		ByteArrayInputStream stream = new ByteArrayInputStream(byteArray);
		
		XMultiComponentFactory xMultiComponentFactory = xComponentContext.getServiceManager();
        Object desktopService = xMultiComponentFactory.createInstanceWithContext("com.sun.star.frame.Desktop", xComponentContext);
        XComponentLoader xComponentLoader = (XComponentLoader) UnoRuntime.queryInterface(XComponentLoader.class, desktopService);
        
        PropertyValue[] propertyValues = new PropertyValue[2];
        propertyValues[0] = new PropertyValue();
        propertyValues[1] = new PropertyValue();

        propertyValues[0].Name = "InputStream";
        propertyValues[0].Value = stream;
        propertyValues[1].Name = "Hidden";
        propertyValues[1].Value = new Boolean(true);

        XComponent xComp = xComponentLoader.loadComponentFromURL("private:stream", "_blank", 0, propertyValues);
        
        XReplaceDescriptor xReplaceDescr = null;
		XReplaceable xReplaceable = null;

		XTextDocument xTextDocument = (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, xComp);

		xReplaceable = (XReplaceable) UnoRuntime.queryInterface(XReplaceable.class, xTextDocument);

		xReplaceDescr = (XReplaceDescriptor) xReplaceable.createReplaceDescriptor();

		  // mail merge the date
		  xReplaceDescr.setSearchString("<date>");
		  xReplaceDescr.setReplaceString(new Date().toString());
		  xReplaceable.replaceAll(xReplaceDescr);
		  
		  // mail merge the addressee
		  xReplaceDescr.setSearchString("<addressee>");
		  xReplaceDescr.setReplaceString("Best Friend");
		  xReplaceable.replaceAll(xReplaceDescr);
		  
		  // mail merge the signatory
		  xReplaceDescr.setSearchString("<signatory>");
		  xReplaceDescr.setReplaceString("Your New Boss");
		  xReplaceable.replaceAll(xReplaceDescr);
		  
		  ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		  
		  propertyValues[0].Name = "OutputStream";
		  propertyValues[0].Value = outputStream;
		  propertyValues[1].Name = "FilterName";
		  propertyValues[1].Value = "writer_pdf_Export";

        XStorable xstorable = (XStorable) UnoRuntime.queryInterface(XStorable.class,xComp);
        xstorable.storeToURL("private:stream", propertyValues);

        XCloseable xclosable = (XCloseable) UnoRuntime.queryInterface(XCloseable.class,xComp);
        xclosable.close(true);

        byte[] outputBytes = outputStream.toByteArray();	  
        
		return outputBytes;
	}
}
