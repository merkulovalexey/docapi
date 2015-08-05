package ru.docapi.customtransformers;


import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map.Entry;

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
	private XComponent template; 
	
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
	
	public void loadTemplate(byte[] bytes) throws Exception {
		
		// transform the steam to LO
		
		OOoInputStream inputStream = new OOoInputStream(bytes);

		
		XMultiComponentFactory xOfficeFactory = xComponentContext.getServiceManager();		
		
		Object desktopService = xOfficeFactory.createInstanceWithContext("com.sun.star.frame.Desktop", xComponentContext);
        XComponentLoader xComponentLoader = (XComponentLoader) UnoRuntime.queryInterface(XComponentLoader.class, desktopService);
        
        PropertyValue[] conversionProperties = new PropertyValue[2];
        conversionProperties[0] = new PropertyValue();
        conversionProperties[1] = new PropertyValue();

        conversionProperties[0].Name = "InputStream";
        conversionProperties[0].Value = inputStream;
        conversionProperties[1].Name = "Hidden";
        conversionProperties[1].Value = new Boolean(true);

        template = xComponentLoader.loadComponentFromURL("private:stream", "_blank", 0, conversionProperties);
		
	}
	
	public void fillTemplate(HashMap<String, String> fields) throws Exception {
		
    
        XReplaceDescriptor xReplaceDescr = null;
		XReplaceable xReplaceable = null;

		XTextDocument xTextDocument = (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, template);

		xReplaceable = (XReplaceable) UnoRuntime.queryInterface(XReplaceable.class, xTextDocument);

		xReplaceDescr = (XReplaceDescriptor) xReplaceable.createReplaceDescriptor();

		for (Entry<String, String> entry : fields.entrySet()) {
					    
		    xReplaceDescr.setSearchString(entry.getKey());
			xReplaceDescr.setReplaceString(entry.getValue());
			xReplaceable.replaceAll(xReplaceDescr);
	
		}       

	}

	
}
